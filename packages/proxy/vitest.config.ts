import { existsSync, readFileSync } from "node:fs";
import { dirname, isAbsolute, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vitest/config";

const packageDir = dirname(fileURLToPath(import.meta.url));

const setupFile = join(packageDir, "test/setup.ts");

const RAW_QUERY = "?raw";

const rawImport = () => ({
  name: "m365-raw-import",
  enforce: "pre" as const,
  load(id: string) {
    if (!id.endsWith(RAW_QUERY)) return null;
    const file = id.slice(0, -RAW_QUERY.length);
    if (!isAbsolute(file) || relative(packageDir, file).startsWith("..")) {
      throw new Error(`raw imports are only read from ${packageDir}, refused ${id}`);
    }
    return { code: `export default ${JSON.stringify(readFileSync(file, "utf8"))};`, moduleSideEffects: false };
  },
});

export default defineConfig({
  plugins: [rawImport()],
  test: {
    root: packageDir,
    include: ["test/**/*.test.ts"],
    setupFiles: [setupFile],
  },
});

if (!existsSync(setupFile)) throw new Error(`missing ${setupFile}`);
