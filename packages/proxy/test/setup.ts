import { existsSync, mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { vi } from "vitest";
import * as h3 from "h3";

const IMPORTS_FILE = ".nitro/types/nitro-imports.d.ts";

const PROXY_PACKAGE_DIR = dirname(dirname(fileURLToPath(import.meta.url)));

const IMPORT_LINE = /^\s+const (\w+): typeof import\('([^']+)'\)\.(\w+)$/gm;

const H3_SOURCE = /\/h3$/;

export const nitroH3Globals = (source: string): string[] => {
  const names = new Set<string>();
  for (const match of source.matchAll(IMPORT_LINE)) {
    if (H3_SOURCE.test(match[2])) names.add(match[3]);
  }
  return [...names].sort();
};

const h3Table = () => h3 as unknown as Record<string, unknown>;

export function installNitroGlobals(packageDir: string): string[] {
  const file = join(packageDir, IMPORTS_FILE);
  if (!existsSync(file)) {
    throw new Error(`${IMPORTS_FILE} is missing; run \`pnpm --filter @m365-copilot/proxy prepare\` before the tests`);
  }
  const names = nitroH3Globals(readFileSync(file, "utf8"));
  for (const name of names) {
    const exported = h3Table()[name];
    if (exported === undefined) throw new Error(`nitro auto-imports ${name}, which h3 does not export`);
    vi.stubGlobal(name, exported);
  }
  vi.stubGlobal("defineNitroConfig", <T>(config: T): T => config);
  vi.stubGlobal("defineNitroPlugin", (plugin: unknown) => plugin);
  return names;
}

const scratch = mkdtempSync(join(tmpdir(), "proxy-msal-"));

process.env.M365_CACHE_FILE = join(scratch, "msal-cache.json");
process.env.M365_SECRETS_FILE = join(scratch, "secrets.json");
delete process.env.M365_ENABLE_INTERACTIVE_APPROVAL;

export interface ConsoleTape {
  logs: string[];
  errors: string[];
  restore(): void;
}

export function recordConsole(): ConsoleTape {
  const logs: string[] = [];
  const errors: string[] = [];
  const log = globalThis.console.log;
  const error = globalThis.console.error;
  globalThis.console.log = (...args: unknown[]) => void logs.push(args.map(String).join(" "));
  globalThis.console.error = (...args: unknown[]) => void errors.push(args.map(String).join(" "));
  return {
    logs,
    errors,
    restore() {
      globalThis.console.log = log;
      globalThis.console.error = error;
    },
  };
}

export const nitroGlobals = installNitroGlobals(PROXY_PACKAGE_DIR);
