import { appendFileSync, mkdirSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { homedir } from "node:os";

const CONFIG_DIR = join(homedir(), ".config", "opencode-m365");
const LOG_FILE = resolve(CONFIG_DIR, process.env.M365_LOG_FILE || "debug.log");
const LOG_DIR = dirname(LOG_FILE);
export const FRAME_DIR = resolve(CONFIG_DIR, process.env.M365_FRAME_DIR || "frames");

const trace = !!process.env.M365_TRACE;
const enabled = !!process.env.M365_DEBUG || trace;
const stdoutEnabled = !!process.env.M365_LOG_STDOUT;

function timestamp(): string {
  return new Date().toISOString();
}

function write(level: string, component: string, ...args: unknown[]) {
  if (!enabled) return;
  const msg = args
    .map((a) => (typeof a === "string" ? a : JSON.stringify(a, null, 2)))
    .join(" ");
  const line = `[${timestamp()}] [${level}] [${component}] ${msg}\n`;
  if (stdoutEnabled) process.stdout.write(line);
  try {
    mkdirSync(LOG_DIR, { recursive: true });
    appendFileSync(LOG_FILE, line);
  } catch {
  }
}

export function createLogger(component: string) {
  return {
    info: (...args: unknown[]) => write("INFO", component, ...args),
    error: (...args: unknown[]) => write("ERROR", component, ...args),
    debug: (...args: unknown[]) => write("DEBUG", component, ...args),
  };
}

export function trunc(value: string, limit: number): string {
  if (trace) return value;
  return value.slice(0, limit);
}

export const LOG_PATH = LOG_FILE;

export function fileLoggingEnabled(): boolean {
  return enabled;
}
