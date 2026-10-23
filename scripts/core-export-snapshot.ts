import { existsSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

export const PACKAGE_NAME = "@m365-copilot/core";

export const DIST_ENTRY = "packages/core/dist/index.mjs";

export const DECLARATION_ENTRY = "packages/core/dist/index.d.mts";

export const DEFAULT_SNAPSHOT_PATH = "packages/core/export-surface.snapshot.json";

export type ExportKind = "value" | "type";

export interface ExportSurface {
  package: string;
  values: string[];
  types: string[];
}

export interface SurfaceDrift {
  addedValues: string[];
  removedValues: string[];
  addedTypes: string[];
  removedTypes: string[];
}

const DECLARATION_EXPORT = /^export\s*\{([^}]*)\}/gm;
const DECLARATION_SPECIFIER = /\s*(type\s+)?([\w$]+)(?:\s+as\s+([\w$]+))?\s*/g;

export function typeNamesFromDeclaration(declaration: string): string[] {
  const names = new Set<string>();
  for (const block of declaration.matchAll(DECLARATION_EXPORT)) {
    for (const specifier of (block[1] ?? "").matchAll(DECLARATION_SPECIFIER)) {
      if (!specifier[1]) continue;
      names.add(specifier[3] || specifier[2] || "");
    }
  }
  return [...names].filter((name) => name.length > 0).sort();
}

const sorted = (names: Iterable<string>) => [...names].sort();

export function buildSurface(packageName: string, runtimeNames: Iterable<string>, declaration: string): ExportSurface {
  return {
    package: packageName,
    values: sorted(runtimeNames),
    types: typeNamesFromDeclaration(declaration),
  };
}

export function surfaceDocument(surface: ExportSurface): string {
  return `${JSON.stringify(surface, null, 2)}\n`;
}

export function compareToSurface(current: ExportSurface, saved: ExportSurface | null): SurfaceDrift {
  if (!saved) {
    return { addedValues: [], removedValues: [], addedTypes: [], removedTypes: [] };
  }
  const added = (next: string[], previous: string[]) => {
    const known = new Set(previous);
    return next.filter((name) => !known.has(name));
  };
  const removed = (next: string[], previous: string[]) => {
    const known = new Set(next);
    return previous.filter((name) => !known.has(name));
  };
  return {
    addedValues: added(current.values, saved.values),
    removedValues: removed(current.values, saved.values),
    addedTypes: added(current.types, saved.types),
    removedTypes: removed(current.types, saved.types),
  };
}

export function driftIsEmpty(drift: SurfaceDrift): boolean {
  return (
    drift.addedValues.length === 0 &&
    drift.removedValues.length === 0 &&
    drift.addedTypes.length === 0 &&
    drift.removedTypes.length === 0
  );
}

export function formatDrift(drift: SurfaceDrift): string[] {
  return [
    ...drift.addedValues.map((name) => `+ value ${name}`),
    ...drift.removedValues.map((name) => `- value ${name}`),
    ...drift.addedTypes.map((name) => `+ type ${name}`),
    ...drift.removedTypes.map((name) => `- type ${name}`),
  ];
}

export function loadSurface(path: string): ExportSurface | null {
  return existsSync(path) ? (JSON.parse(readFileSync(path, "utf8")) as ExportSurface) : null;
}

export function readDeclaration(path: string): string {
  if (!existsSync(path)) throw new Error(`${path} is missing; run pnpm build first`);
  return readFileSync(path, "utf8");
}

export async function importSurface(entry: string): Promise<string[]> {
  if (!existsSync(entry)) throw new Error(`${entry} is missing; run pnpm build first`);
  const exported: Record<string, unknown> = await import(pathToFileURL(entry).href);
  return Object.keys(exported);
}

export interface RunOptions {
  repoRoot?: string;
  stdout?: (line: string) => void;
}

export type SnapshotMode = "check" | "write";

export function parseMode(argv: string[]): SnapshotMode {
  const flag = argv.find((arg) => arg.startsWith("--"));
  if (flag === "--write") return "write";
  if (flag === undefined || flag === "--check") return "check";
  throw new Error(`unknown flag "${flag}"; use --check or --write`);
}

export async function run(argv: string[], options: RunOptions = {}): Promise<number> {
  const repoRoot = options.repoRoot ?? process.cwd();
  const emit = options.stdout ?? ((line: string) => process.stdout.write(`${line}\n`));
  const mode = parseMode(argv);
  const snapshotPath = join(repoRoot, DEFAULT_SNAPSHOT_PATH);
  const current = buildSurface(
    PACKAGE_NAME,
    await importSurface(join(repoRoot, DIST_ENTRY)),
    readDeclaration(join(repoRoot, DECLARATION_ENTRY)),
  );
  const counts = `${current.values.length} values, ${current.types.length} types`;
  if (mode === "write") {
    writeFileSync(snapshotPath, surfaceDocument(current));
    emit(`wrote ${DEFAULT_SNAPSHOT_PATH}: ${counts}`);
    return 0;
  }
  const saved = loadSurface(snapshotPath);
  if (!saved) {
    emit(`snapshot ${DEFAULT_SNAPSHOT_PATH} is missing; run pnpm snapshot:exports`);
    return 1;
  }
  const drift = formatDrift(compareToSurface(current, saved));
  if (drift.length === 0) {
    emit(`export surface matches ${DEFAULT_SNAPSHOT_PATH}: ${counts}`);
    return 0;
  }
  emit(...drift);
  emit(`export surface drifted from ${DEFAULT_SNAPSHOT_PATH}: ${counts}`);
  return 1;
}

export const invokedDirectly =
  process.argv[1] !== undefined &&
  pathToFileURL(resolve(process.argv[1])).href === pathToFileURL(fileURLToPath(import.meta.url)).href;

if (invokedDirectly) process.exitCode = await run(process.argv.slice(2));
