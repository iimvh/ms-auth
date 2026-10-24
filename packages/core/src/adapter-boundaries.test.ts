import { afterEach, describe, expect, it } from "vitest";
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import {
  ADAPTER_ROOT,
  buildGraph,
  checkBoundaries,
  compareToBaseline,
  findCrossAdapterImports,
  loadBaseline,
  run,
} from "../../../scripts/adapter-boundaries.js";

const CORE_SRC = "packages/core/src";

const SHARED_MODULE = "export const shared = 1;\nexport const other = 1;\n";
const MSAL_IMPORTS_PERSONAL = "import { loadPersonal } from \"../personal/personal.js\";\n";
const PERSONAL_IMPORTS_MSAL = "import { getToken } from \"../msal/token.js\";\n";
const SAME_ADAPTER_IMPORT = "import { other } from \"./other.js\";\n";
const SHARED_IMPORT = "import { shared } from \"../shared.js\";\n";

const scratchDirs: string[] = [];

function fixtureAt(root: string, files: Record<string, string>): void {
  for (const [path, contents] of Object.entries(files)) {
    const full = path.startsWith(`${ADAPTER_ROOT}/`) ? join(root, path) : join(root, CORE_SRC, path);
    mkdirSync(dirname(full), { recursive: true });
    writeFileSync(full, contents);
  }
}

function fixture(files: Record<string, string>): string {
  const root = mkdtempSync(join(tmpdir(), "adapter-boundaries-"));
  scratchDirs.push(root);
  fixtureAt(root, files);
  return root;
}

const silent = { stdout: () => {} };

function crossImports() {
  return fixture({
    "shared.ts": SHARED_MODULE,
    "index.ts": "export { shared } from \"./shared.js\";\n",
    [`${ADAPTER_ROOT}/msal/token.ts`]: MSAL_IMPORTS_PERSONAL,
    [`${ADAPTER_ROOT}/msal/other.ts`]: "export const other = 1;\n",
    [`${ADAPTER_ROOT}/msal/login.ts`]: `${SAME_ADAPTER_IMPORT}${SHARED_IMPORT}`,
    [`${ADAPTER_ROOT}/personal/personal.ts`]: `${SAME_ADAPTER_IMPORT}${SHARED_IMPORT}${PERSONAL_IMPORTS_MSAL}`,
  });
}

afterEach(() => {
  while (scratchDirs.length) rmSync(scratchDirs.pop()!, { recursive: true, force: true });
});

describe("adapter boundary", () => {
  it("passes on the current packages/core tree, with both adapters present", () => {
    const report = checkBoundaries({ coreSrcDir: join(process.cwd(), CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.adaptersPresent).toEqual(["msal", "personal"]);
    expect(report.scannedFiles).toBeGreaterThan(20);
    expect(report.violations).toEqual([]);
  });

  it("fails --check on a copy of the current graph once an adapter reaches into the other", () => {
    const root = mkdtempSync(join(tmpdir(), "adapter-boundaries-"));
    scratchDirs.push(root);
    const graph = buildGraph({ adapterRoot: ADAPTER_ROOT });
    expect(graph.edges.some((edge) => edge.from === `${ADAPTER_ROOT}/msal/auth.ts`)).toBe(true);
    for (const node of graph.nodes.filter((node) => node.startsWith(`${CORE_SRC}/`))) {
      const full = join(root, node);
      mkdirSync(dirname(full), { recursive: true });
      writeFileSync(full, readFileSync(node, "utf8"));
    }
    fixtureAt(root, {
      [`${ADAPTER_ROOT}/msal/token.ts`]: MSAL_IMPORTS_PERSONAL,
      [`${ADAPTER_ROOT}/personal/personal.ts`]: `export const personal = 1;\n`,
    });
    const report = checkBoundaries({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.violations.map((v) => `${v.importer}:${v.adapter}->${v.targetAdapter}`)).toEqual([
      `${ADAPTER_ROOT}/msal/token.ts:msal->personal`,
    ]);
    expect(run(["--check"], { repoRoot: root, ...silent })).toBe(1);
  });

  it("reports an msal import of personal and the reverse", () => {
    const root = crossImports();
    const report = checkBoundaries({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.adaptersPresent).toEqual(["msal", "personal"]);
    expect(report.violations.map((v) => `${v.importer}:${v.adapter}->${v.targetAdapter}`)).toEqual([
      `${ADAPTER_ROOT}/msal/token.ts:msal->personal`,
      `${ADAPTER_ROOT}/personal/personal.ts:personal->msal`,
    ]);
    expect(report.violations.map((v) => v.rule)).toEqual(["adapter-isolation", "adapter-isolation"]);
  });

  it("does not report a same-adapter import or an import of a shared module", () => {
    const root = crossImports();
    const violations = findCrossAdapterImports(buildGraph({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT }));
    expect(violations.map((v) => v.importer)).not.toContain(`${ADAPTER_ROOT}/msal/login.ts`);
    expect(violations.filter((v) => v.specifier === "../shared.js")).toEqual([]);
    expect(violations).toHaveLength(2);
  });

  it("reports a violation written as a dynamic import", () => {
    const root = fixture({
      [`${ADAPTER_ROOT}/msal/lazy.ts`]: "export const load = () => import(\"../personal/personal.js\");\n",
      [`${ADAPTER_ROOT}/personal/personal.ts`]: "export const personal = 1;\n",
    });
    const report = checkBoundaries({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.violations.map((v) => `${v.importer}:${v.adapter}->${v.targetAdapter}`)).toEqual([
      `${ADAPTER_ROOT}/msal/lazy.ts:msal->personal`,
    ]);
  });

  it("reports a violation whose import clause spans several lines", () => {
    const root = fixture({
      [`${ADAPTER_ROOT}/msal/wide.ts`]: "import {\n  personal,\n  other,\n} from \"../personal/personal.js\";\nexport const both = [personal, other];\n",
      [`${ADAPTER_ROOT}/personal/personal.ts`]: "export const personal = 1;\nexport const other = 1;\n",
    });
    const report = checkBoundaries({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.violations.map((v) => `${v.importer}:${v.adapter}->${v.targetAdapter}`)).toEqual([
      `${ADAPTER_ROOT}/msal/wide.ts:msal->personal`,
    ]);
  });

  it("passes when only one adapter exists", () => {    const root = fixture({ [`${ADAPTER_ROOT}/msal/token.ts`]: SHARED_IMPORT });
    const report = checkBoundaries({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    expect(report.adaptersPresent).toEqual(["msal"]);
    expect(report.violations).toEqual([]);
  });

  it("exits 1 from --check when an adapter crosses over, and 0 once the import is gone", () => {
    const root = crossImports();
    expect(run(["--check"], { repoRoot: root, ...silent })).toBe(1);
    expect(run(["--write-baseline"], { repoRoot: root, ...silent })).toBe(0);
    expect(run(["--check"], { repoRoot: root, ...silent })).toBe(1);
  });

  it("exits 1 when an adapter reaches the other through a repo-root path", () => {
    const root = fixture({
      [`${ADAPTER_ROOT}/msal/token.ts`]: `import { personal } from "packages/core/src/adapters/personal/personal.js";\n`,
      [`${ADAPTER_ROOT}/personal/personal.ts`]: "export const personal = 1;\n",
    });
    expect(run(["--check"], { repoRoot: root, ...silent })).toBe(1);
  });

  it("catches a cross-import written from a subdirectory at the same depth as an adapter", () => {
    const root = fixture({
      "zz-probe/probe.ts": `import { personal } from "../adapters/personal/personal.js";\n`,
      [`${ADAPTER_ROOT}/personal/personal.ts`]: "export const personal = 1;\n",
      [`${ADAPTER_ROOT}/msal/token.ts`]: "export const token = 1;\n",
    });
    const graph = buildGraph({ coreSrcDir: join(root, CORE_SRC), adapterRoot: ADAPTER_ROOT });
    const probe = graph.edges.filter((edge) => edge.from === `${CORE_SRC}/zz-probe/probe.ts`);
    expect(probe.map((edge) => edge.to)).toEqual([`${ADAPTER_ROOT}/personal/personal.ts`]);
  });
});

describe("dependency graph baseline", () => {
  it("keeps the transport free of every auth, personal and account module (#9)", () => {
    const graph = buildGraph({ adapterRoot: ADAPTER_ROOT });
    const transport = `${CORE_SRC}/protocol/session.ts`;
    const reach = graph.edges.filter((edge) => edge.from === transport).map((edge) => edge.to);
    expect(reach.filter((to) => /adapters\/|account-registry|personal-chathub|personal-wiring/.test(to))).toEqual([]);
    expect(reach).toContain(`${CORE_SRC}/turn-context.ts`);
  });

  it("matches the committed baseline and reports no drift", () => {
    const graph = buildGraph({ adapterRoot: ADAPTER_ROOT });
    const drift = compareToBaseline(graph, loadBaseline());
    expect(drift.missingBaseline).toBe(false);
    expect(drift.addedEdges).toEqual([]);
    expect(drift.removedEdges).toEqual([]);
    expect(drift.addedNodes).toEqual([]);
    expect(drift.removedNodes).toEqual([]);
  });

  it("reports an edge the baseline is missing and an edge the tree dropped", () => {
    const graph = buildGraph({ adapterRoot: ADAPTER_ROOT });
    const saved = loadBaseline()!;
    const first = saved.edges[0];
    const extra = { from: `${CORE_SRC}/ghost.ts`, to: `${CORE_SRC}/protocol/session.ts`, specifier: "./session.js", line: 1 };
    const drift = compareToBaseline(graph, { ...saved, edges: [...saved.edges.slice(1), extra] });
    expect(drift.addedEdges.map((e) => `${e.from}->${e.to}`)).toEqual([`${first.from}->${first.to}`]);
    expect(drift.removedEdges).toEqual([extra]);
    expect(drift.addedNodes).toEqual([]);
    expect(drift.removedNodes).toEqual([]);
  });

  it("reports a module the baseline does not know", () => {
    const graph = buildGraph({ adapterRoot: ADAPTER_ROOT });
    const saved = loadBaseline()!;
    const drift = compareToBaseline(graph, { ...saved, nodes: saved.nodes.filter((node) => !node.endsWith("/session.ts")) });
    expect(drift.addedNodes).toContain(`${CORE_SRC}/protocol/session.ts`);
    expect(drift.removedNodes).toEqual([]);
  });

  it("treats a missing baseline as drift", () => {
    expect(compareToBaseline(buildGraph({ adapterRoot: ADAPTER_ROOT }), null).missingBaseline).toBe(true);
  });

  it("commits the baseline as data covering every core module and its edges", () => {
    const baseline = JSON.parse(readFileSync("scripts/dependency-graph-baseline.json", "utf8"));
    expect(baseline.nodes).toContain(`${CORE_SRC}/index.ts`);
    expect(baseline.nodes).toContain(`${CORE_SRC}/protocol/session.ts`);
    expect(baseline.nodes).toContain(`${ADAPTER_ROOT}/msal/auth.ts`);
    const reaches = (from: string, to: string) =>
      baseline.edges.some((e: { from: string; to: string }) => e.from === from && e.to === to);
    // The credential decision lives in the app; the transport is handed a context
    // and reaches no adapter, auth or account module at all (#9).
    expect(reaches(`${CORE_SRC}/app/credential.ts`, `${CORE_SRC}/adapters/msal/auth.ts`)).toBe(true);
    expect(reaches(`${CORE_SRC}/app/credential.ts`, `${CORE_SRC}/adapters/personal/personal-auth.ts`)).toBe(true);
    for (const adapter of [`${ADAPTER_ROOT}/msal/auth.ts`, `${ADAPTER_ROOT}/personal/personal-auth.ts`, `${CORE_SRC}/account-registry.ts`]) {
      expect(reaches(`${CORE_SRC}/protocol/session.ts`, adapter)).toBe(false);
    }
    expect(baseline.nodes.length).toBeGreaterThan(20);
  });
});

describe("lint:boundaries wiring", () => {
  it("is registered in the root package.json and runs green on the current tree", () => {
    const rootPackage = JSON.parse(readFileSync("package.json", "utf8"));
    expect(rootPackage.scripts["lint:boundaries"]).toBe("node scripts/adapter-boundaries.ts --check");
    expect(rootPackage.scripts["lint:boundaries:baseline"]).toBe("node scripts/adapter-boundaries.ts --write-baseline");
    const result = spawnSync(process.execPath, ["scripts/adapter-boundaries.ts", "--check"], { encoding: "utf8" });
    expect(result.status).toBe(0);
    expect(result.stdout).toContain("boundary check passed");
  });
});
