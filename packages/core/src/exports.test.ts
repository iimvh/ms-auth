import { afterEach, describe, expect, it } from "vitest";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  DECLARATION_ENTRY,
  DEFAULT_SNAPSHOT_PATH,
  buildSurface,
  compareToSurface,
  driftIsEmpty,
  formatDrift,
  loadSurface,
  readDeclaration,
  run,
  typeNamesFromDeclaration,
  type ExportSurface,
} from "../../../scripts/core-export-snapshot.js";

const REPO_ROOT = process.cwd();

const PACKAGE = "@m365-copilot/core";

const DECLARATION = "export {\n  value,\n  type Shape,\n  Alias as Renamed,\n  type Inner as Outer,\n};\n";

const RUNTIME = ["value", "Alias"];

const scratchDirs: string[] = [];

function writeSnapshotAt(contents: string): string {
  const root = mkdtempSync(join(tmpdir(), "export-surface-"));
  scratchDirs.push(root);
  const path = join(root, "snapshot.json");
  writeFileSync(path, contents);
  return path;
}

function savedSurface(): ExportSurface {
  const saved = loadSurface(join(REPO_ROOT, DEFAULT_SNAPSHOT_PATH));
  if (!saved) throw new Error(`${DEFAULT_SNAPSHOT_PATH} is missing; run pnpm snapshot:exports`);
  return saved;
}

function expectedCounts(): string {
  const saved = savedSurface();
  return `${saved.values.length} values, ${saved.types.length} types`;
}

afterEach(() => {
  while (scratchDirs.length) rmSync(scratchDirs.pop()!, { recursive: true, force: true });
});

describe("typeNamesFromDeclaration", () => {
  it("reads every type-only specifier and keeps its local alias", () => {
    expect(typeNamesFromDeclaration(DECLARATION)).toEqual(["Outer", "Shape"]);
  });

  it("scans an export clause spread across lines", () => {
    expect(typeNamesFromDeclaration("export {\n  value,\n  type First,\n  type Second,\n};\n")).toEqual(["First", "Second"]);
  });
});

describe("compareToSurface", () => {
  const saved = buildSurface(PACKAGE, RUNTIME, DECLARATION);

  it("reports nothing when the surface is unchanged", () => {
    expect(driftIsEmpty(compareToSurface(buildSurface(PACKAGE, RUNTIME, DECLARATION), saved))).toBe(true);
  });

  it("reports an added export of either kind", () => {
    const current = buildSurface(PACKAGE, [...RUNTIME, "extra"], `export { type Extra };\n${DECLARATION}`);
    expect(formatDrift(compareToSurface(current, saved))).toEqual(["+ value extra", "+ type Extra"]);
  });

  it("reports a removed or renamed export of either kind", () => {
    const current = buildSurface(PACKAGE, ["value"], "export { value, type Shape, type Inner as Outer };\n");
    expect(formatDrift(compareToSurface(current, saved))).toEqual(["- value Alias"]);
    const renamed = buildSurface(PACKAGE, RUNTIME, DECLARATION.replace("Outer", "Outermost"));
    expect(formatDrift(compareToSurface(renamed, saved))).toEqual(["+ type Outermost", "- type Outer"]);
  });
});

describe("export surface of @m365-copilot/core", () => {
  it("matches the committed snapshot", async () => {
    const lines: string[] = [];
    await run(["--check"], { repoRoot: REPO_ROOT, stdout: (line) => lines.push(line) });
    expect(lines).toEqual([`export surface matches ${DEFAULT_SNAPSHOT_PATH}: ${expectedCounts()}`]);
  }, 30_000);

  it("agrees with a fresh scan of the built declaration", () => {
    expect(typeNamesFromDeclaration(readDeclaration(join(REPO_ROOT, DECLARATION_ENTRY)))).toEqual(savedSurface().types);
  });

  it("flags a value the snapshot does not yet list", () => {
    const saved = savedSurface();
    const trailingValue = saved.values[saved.values.length - 1];
    const stale = loadSurface(writeSnapshotAt(JSON.stringify({ ...saved, values: saved.values.filter((name) => name !== trailingValue) })));
    expect(formatDrift(compareToSurface(saved, stale))).toEqual([`+ value ${trailingValue}`]);
  });

  it("flags a type the snapshot lists but the build no longer declares", () => {
    const saved = savedSurface();
    const stale = loadSurface(writeSnapshotAt(JSON.stringify({ ...saved, types: [...saved.types, "Ghost"] })));
    expect(formatDrift(compareToSurface(saved, stale))).toEqual(["- type Ghost"]);
  });
});
