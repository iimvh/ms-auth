import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { runPersonalCli } from "./personal-cli.js";
import { PersonalStore } from "./personal-session.js";

const capture = {
  accessToken: "secret-token", chathubPath: "oid@tid", scenario: "OfficeWebPaidConsumerCopilot",
  variants: "a,b", licenseType: "Consumer", agentHost: "Bizchat.FullScreen", product: "Office",
  source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
  expiresAt: 9_999_999_999_999, capturedAt: 0,
};
const deps = () => {
  const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
  const out: string[] = [];
  return { store, out, capture: vi.fn(async () => capture), probe: vi.fn(async () => ({ model: "claude-sonnet", ok: true, at: 1 })), outFn: (l: string) => out.push(l), now: () => 1 };
};

describe("m365-personal cli", () => {
  it("login writes the record and reports the probe, never the token", async () => {
    const d = deps();
    const code = await runPersonalCli(["login"], { store: d.store, capture: d.capture, probe: d.probe, out: d.outFn, now: d.now });
    expect(code).toBe(0);
    expect(d.store.read()?.mintCount).toBe(1);
    expect(d.store.read()?.probe?.model).toBe("claude-sonnet");
    expect(d.out.join("\n")).not.toContain("secret-token");
  });

  it("status prints expiry, scenario and path but not the token; --json is machine-readable", async () => {
    const d = deps();
    await runPersonalCli(["login"], { store: d.store, capture: d.capture, probe: d.probe, out: d.outFn, now: d.now });
    d.out.length = 0;
    expect(await runPersonalCli(["status"], { store: d.store, out: d.outFn, now: d.now })).toBe(0);
    expect(d.out.join("\n")).toContain("OfficeWebPaidConsumerCopilot");
    expect(d.out.join("\n")).not.toContain("secret-token");
    d.out.length = 0;
    await runPersonalCli(["status", "--json"], { store: d.store, out: d.outFn, now: d.now });
    const parsed = JSON.parse(d.out.join("\n"));
    expect(parsed.scenario).toBe("OfficeWebPaidConsumerCopilot");
    expect(parsed.accessToken).toBeUndefined();
  });

  it("login exits non-zero when no candidate model answers (never reports unusable capture as success)", async () => {
    const d = deps();
    d.probe = vi.fn(async () => ({ model: "claude-sonnet", ok: false, at: 1 }));
    const code = await runPersonalCli(["login"], { store: d.store, capture: d.capture, probe: d.probe, out: d.outFn, now: d.now });
    expect(code).not.toBe(0);
    expect(d.out.join("\n")).toContain("does not serve this account");
    // the capture is still recorded: a working probe model may exist after a re-login
    expect(d.store.read()?.scenario).toBe("OfficeWebPaidConsumerCopilot");
  });

  it("status without a record exits 1 and says to log in", async () => {
    const d = deps();
    expect(await runPersonalCli(["status"], { store: d.store, out: d.outFn, now: d.now })).toBe(1);
    expect(d.out.join("\n")).toContain("pnpm m365-personal login");
  });

  it("logout deletes, an unknown command exits 2", async () => {
    const d = deps();
    await runPersonalCli(["login"], { store: d.store, capture: d.capture, probe: d.probe, out: d.outFn, now: d.now });
    expect(await runPersonalCli(["logout"], { store: d.store, out: d.outFn })).toBe(0);
    expect(d.store.read()).toBeNull();
    expect(await runPersonalCli(["wat"], { store: d.store, out: d.outFn })).toBe(2);
  });
});
