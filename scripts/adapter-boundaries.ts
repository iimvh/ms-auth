import { existsSync, mkdirSync, readFileSync, readdirSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve, sep } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

export const ADAPTER_ROOT = "packages/core/src/adapters";

export const ADAPTER_NAMES = ["msal", "personal"] as const;

export const CROSS_ADAPTER_RULE = "adapter-isolation";

export const DEFAULT_BASELINE_PATH = "scripts/dependency-graph-baseline.json";

export type AdapterName = (typeof ADAPTER_NAMES)[number];
export type BaselineMode = "create" | "check";

export interface ImportEdge {
  from: string;
  to: string;
  specifier: string;
  line: number;
}

export interface GraphOptions {
  coreSrcDir?: string;
  adapterRoot?: string;
}

export interface BoundaryViolation {
  rule: typeof CROSS_ADAPTER_RULE;
  importer: string;
  adapter: AdapterName;
  targetAdapter: AdapterName;
  specifier: string;
  line: number;
}

export interface BoundaryReport {
  adapterRoot: string;
  adaptersPresent: AdapterName[];
  scannedFiles: number;
  violations: BoundaryViolation[];
}

export interface DependencyGraph {
  adapterRoot: string;
  nodes: string[];
  edges: ImportEdge[];
}

export interface GraphDrift {
  addedEdges: ImportEdge[];
  removedEdges: ImportEdge[];
  addedNodes: string[];
  removedNodes: string[];
  missingBaseline: boolean;
}

const CODE_FILE = /\.(?:ts|mts|tsx|js|mjs)$/;
const SPECIFIER = /^(?!\s*(?:\/\/|\/\*|\*))[ \t]*(?:(?:import|export)\b[^;'"*\n]*?\bfrom[ \t]*|(?:const|let|var)[ \t]+[^=;*\n]*?=[ \t]*require[ \t]*\([ \t]*)(["'])([^"'\n]+)\1|\b(?:import|require)[ \t]*\([ \t]*(["'])([^"'\n]+)\3/gm;
const RESOLVE_SUFFIXES = ["", ".ts", ".tsx", ".mts", ".js", ".mjs", "/index.ts", "/index.js", "/index.mjs"];

const toPosix = (p: string) => p.split(sep).join("/");

const repoRoot = () => process.cwd();

const coreSrcDirOf = (options: GraphOptions) => options.coreSrcDir ?? join(repoRoot(), "packages/core/src");

const repoRootOf = (coreSrcDir: string) => resolve(coreSrcDir, "../../..");

export function coreFiles(srcDir: string): string[] {
  if (!existsSync(srcDir)) return [];
  const found: string[] = [];
  const walk = (dir: string) => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (entry.name === "node_modules" || entry.name.startsWith(".")) continue;
      const full = join(dir, entry.name);
      if (entry.isDirectory()) walk(full);
      else if (CODE_FILE.test(entry.name) && !/\.test\.[cm]?[jt]sx?$/.test(entry.name) && !/\.d\.[cm]?ts$/.test(entry.name)) found.push(full);
    }
  };
  walk(srcDir);
  return found.sort();
}

export function importedSpecifiers(source: string): { specifier: string; line: number }[] {
  return [...source.matchAll(SPECIFIER)].map((match) => ({
    specifier: match[2] ?? match[4],
    line: source.slice(0, match.index).split("\n").length,
  }));
}

function candidatePaths(file: string, specifier: string): string[] {
  const base = resolve(dirname(file), specifier);
  if (!base.endsWith(".js")) return RESOLVE_SUFFIXES.map((suffix) => base + suffix);
  const stem = base.slice(0, -3);
  return [base, `${stem}.ts`, `${stem}.tsx`, `${stem}/index.ts`];
}

function resolveSpecifier(file: string, specifier: string): string | null {
  if (!specifier.startsWith(".")) return null;
  return candidatePaths(file, specifier).find((candidate) => existsSync(candidate)) ?? null;
}

export function importEdges(file: string, repoRoot: string): ImportEdge[] {
  const from = toPosix(relative(repoRoot, file));
  return importedSpecifiers(readFileSync(file, "utf8")).map(({ specifier, line }) => {
    const target = resolveSpecifier(file, specifier);
    return { from, to: target ? toPosix(relative(repoRoot, target)) : specifier, specifier, line };
  });
}

function compareEdges(a: ImportEdge, b: ImportEdge): number {
  return a.from.localeCompare(b.from) || a.to.localeCompare(b.to) || a.specifier.localeCompare(b.specifier);
}

export function buildGraph(options: GraphOptions = {}): DependencyGraph {
  const srcDir = coreSrcDirOf(options);
  const repoRoot = repoRootOf(srcDir);
  const edges = coreFiles(srcDir).flatMap((file) => importEdges(file, repoRoot)).sort(compareEdges);
  return {
    adapterRoot: options.adapterRoot ?? ADAPTER_ROOT,
    nodes: [...new Set(edges.flatMap((edge) => [edge.from, edge.to]))].sort(),
    edges,
  };
}

export function adapterOf(path: string, adapterRoot: string = ADAPTER_ROOT): AdapterName | null {
  for (const name of ADAPTER_NAMES) {
    if (path === `${adapterRoot}/${name}` || path.startsWith(`${adapterRoot}/${name}/`)) return name;
  }
  return null;
}

export function findCrossAdapterImports(graph: DependencyGraph): BoundaryViolation[] {
  const violations: BoundaryViolation[] = [];
  for (const edge of graph.edges) {
    const adapter = adapterOf(edge.from, graph.adapterRoot);
    const targetAdapter = adapterOf(edge.to, graph.adapterRoot);
    if (!adapter || !targetAdapter || adapter === targetAdapter) continue;
    violations.push({ rule: CROSS_ADAPTER_RULE, importer: edge.from, adapter, targetAdapter, specifier: edge.specifier, line: edge.line });
  }
  return violations;
}

function presentAdapters(srcDir: string): AdapterName[] {
  return ADAPTER_NAMES.filter((name) => existsSync(join(srcDir, "adapters", name)));
}

export function checkBoundaries(options: GraphOptions = {}): BoundaryReport {
  const srcDir = coreSrcDirOf(options);
  const graph = buildGraph(options);
  return {
    adapterRoot: graph.adapterRoot,
    adaptersPresent: presentAdapters(srcDir),
    scannedFiles: coreFiles(srcDir).length,
    violations: findCrossAdapterImports(graph),
  };
}

export function formatViolation(violation: BoundaryViolation): string {
  return `${violation.importer}:${violation.line} imports ${violation.specifier} — adapters/${violation.adapter} must not import adapters/${violation.targetAdapter} (${violation.rule})`;
}

function omitRelativeSpecifier(this: unknown, key: string, value: unknown) {
  return key === "specifier" && typeof value === "string" && value.startsWith("./") ? undefined : value;
}

export function baselineDocument(graph: DependencyGraph): string {
  return `${JSON.stringify({ adapterRoot: graph.adapterRoot, nodes: graph.nodes, edges: graph.edges }, omitRelativeSpecifier, 2)}\n`;
}

const edgeKey = (edge: ImportEdge) => `${edge.from}\u0000${edge.to}`;

export function compareToBaseline(graph: DependencyGraph, baseline: DependencyGraph | null): GraphDrift {
  if (!baseline) return { addedEdges: [], removedEdges: [], addedNodes: [], removedNodes: [], missingBaseline: true };
  const savedEdges = new Set(baseline.edges.map(edgeKey));
  const currentEdges = new Set(graph.edges.map(edgeKey));
  const savedNodes = new Set(baseline.nodes);
  const currentNodes = new Set(graph.nodes);
  return {
    addedEdges: graph.edges.filter((edge) => !savedEdges.has(edgeKey(edge))),
    removedEdges: baseline.edges.filter((edge) => !currentEdges.has(edgeKey(edge))),
    addedNodes: graph.nodes.filter((node) => !savedNodes.has(node)),
    removedNodes: baseline.nodes.filter((node) => !currentNodes.has(node)),
    missingBaseline: false,
  };
}

export function loadBaseline(path: string = DEFAULT_BASELINE_PATH): DependencyGraph | null {
  return existsSync(path) ? (JSON.parse(readFileSync(path, "utf8")) as DependencyGraph) : null;
}

function writeBaseline(graph: DependencyGraph, path: string = DEFAULT_BASELINE_PATH): string {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, baselineDocument(graph));
  return `wrote ${path}: ${graph.nodes.length} nodes, ${graph.edges.length} edges`;
}

function parseMode(argv: string[]): BaselineMode {
  const flag = argv.find((arg) => arg.startsWith("--"));
  if (flag === "--write-baseline") return "create";
  if (flag === undefined || flag === "--check") return "check";
  throw new Error(`unknown flag "${flag}"; use --check or --write-baseline`);
}

export interface RunOptions {
  repoRoot?: string;
  stdout?: (line: string) => void;
}

export function run(argv: string[], options: RunOptions = {}): number {
  const repoRoot = options.repoRoot ?? process.cwd();
  const emit = options.stdout ?? ((line: string) => process.stdout.write(`${line}\n`));
  const mode = parseMode(argv);
  const graphOptions: GraphOptions = { coreSrcDir: join(repoRoot, "packages/core/src"), adapterRoot: ADAPTER_ROOT };
  const baselinePath = join(repoRoot, DEFAULT_BASELINE_PATH);
  const report = checkBoundaries(graphOptions);
  const graph = buildGraph(graphOptions);
  const lines = [`adapters present: ${report.adaptersPresent.join(", ") || "none"}`, `scanned ${report.scannedFiles} files under packages/core/src`];
  if (mode === "create") {
    lines.unshift(writeBaseline(graph, baselinePath), ...report.violations.map(formatViolation));
    lines.forEach(emit);
    return 0;
  }
  const drift = compareToBaseline(graph, loadBaseline(baselinePath));
  lines.push(...report.violations.map(formatViolation));
  if (drift.missingBaseline) lines.push(`baseline ${DEFAULT_BASELINE_PATH} is missing; run pnpm lint:boundaries:baseline`);
  lines.push(...drift.addedEdges.map((edge) => `graph drift + ${edge.from} -> ${edge.to}`));
  lines.push(...drift.removedEdges.map((edge) => `graph drift - ${edge.from} -> ${edge.to}`));
  lines.push(...drift.addedNodes.map((node) => `graph drift + node ${node}`));
  lines.push(...drift.removedNodes.map((node) => `graph drift - node ${node}`));
  const failed = report.violations.length > 0 || drift.missingBaseline;
  lines.push(failed ? "boundary check failed" : "boundary check passed: 0 cross-adapter imports");
  lines.forEach(emit);
  return failed ? 1 : 0;
}

export const invokedDirectly = process.argv[1] !== undefined && pathToFileURL(resolve(process.argv[1])).href === pathToFileURL(fileURLToPath(import.meta.url)).href;

if (invokedDirectly) process.exitCode = run(process.argv.slice(2));
