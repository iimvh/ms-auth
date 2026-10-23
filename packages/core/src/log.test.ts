import { afterEach, describe, expect, it, vi } from "vitest";
import { homedir } from "node:os";
import { join, parse, resolve } from "node:path";

const CONFIG_DIR = join(homedir(), ".config", "opencode-m365");
const ABSOLUTE_LOG = join(parse(CONFIG_DIR).root, "tmp", "x", "a.log");

/** log.ts reads the env once, at load, like the proxy does at startup. */
async function load(env: Record<string, string | undefined>) {
  for (const [k, v] of Object.entries(env)) vi.stubEnv(k, v);
  vi.resetModules();
  return import("./log.js");
}

describe("debug log and frame dump paths", () => {
  afterEach(() => vi.unstubAllEnvs());

  it("default to debug.log and frames/ in the config dir", async () => {
    const m = await load({ M365_LOG_FILE: undefined, M365_FRAME_DIR: undefined });
    expect(m.LOG_PATH).toBe(join(CONFIG_DIR, "debug.log"));
    expect(m.FRAME_DIR).toBe(join(CONFIG_DIR, "frames"));
  });

  it("take M365_LOG_FILE and M365_FRAME_DIR relative to the config dir", async () => {
    const m = await load({ M365_LOG_FILE: "my/log.log", M365_FRAME_DIR: "my/frames" });
    expect(m.LOG_PATH).toBe(join(CONFIG_DIR, "my", "log.log"));
    expect(m.FRAME_DIR).toBe(join(CONFIG_DIR, "my", "frames"));
  });

  it("take an absolute path as is, and an empty one as unset", async () => {
    const m = await load({ M365_LOG_FILE: ABSOLUTE_LOG, M365_FRAME_DIR: "" });
    expect(m.LOG_PATH).toBe(resolve(ABSOLUTE_LOG));
    expect(m.FRAME_DIR).toBe(join(CONFIG_DIR, "frames"));
  });
});
