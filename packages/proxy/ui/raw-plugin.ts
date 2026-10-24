import { readFileSync } from "node:fs";
import { isAbsolute, relative } from "node:path";

export const RAW_QUERY = "?raw";

// Nitro's build makes readFileSync unreliable; a platform constraint.

export interface RawModule {
  code: string;
  moduleSideEffects: false;
}

export interface RawPlugin {
  name: string;
  enforce: "pre";
  load(id: string): RawModule | null;
}

export const modulePath = (id: string): string => (id.endsWith(RAW_QUERY) ? id.slice(0, -RAW_QUERY.length) : id);

export function rawImport(rootDir: string = process.cwd()): RawPlugin {
  return {
    name: "m365-raw-import",
    enforce: "pre",
    load(id: string): RawModule | null {
      if (!id.endsWith(RAW_QUERY)) return null;
      const file = modulePath(id);
      if (!isAbsolute(file) || relative(rootDir, file).startsWith("..")) {
        throw new Error(`raw imports are only read from ${rootDir}, refused ${id}`);
      }
      return {
        code: `export default ${JSON.stringify(readFileSync(file, "utf8"))};`,
        moduleSideEffects: false,
      };
    },
  };
}
