import { existsSync, mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { runPersonalCli } from "./personal-cli.js";
import { PersonalStore } from "./personal-session.js";
import { accountsDir, activateAccount, clearMockRegistry, configDir, useMockRegistry } from "./account-registry.js";

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

  it("logout refuses a profileDir that is the data dir itself", async () => {
    // `logout` is an rm -rf of profileDir, and a hand-edited registry entry can
    // point it at the data dir, which would take secrets.json and every other
    // account with it.
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }]);
    try {
      const out: string[] = [];
      let err: { code?: string; message: string } | null = null;
      try {
        await runPersonalCli(["logout"], {
          store: new PersonalStore(join(configDir(), "personal-session.json")),
          profileDir: configDir(),
          out: (l: string) => out.push(l),
        });
      } catch (e) { err = e as { code?: string; message: string }; }
      expect(err).not.toBeNull();
      expect(err!.code).toBe("registry_invalid");
      expect(err!.message).toContain("refusing to delete");
      expect(existsSync(configDir())).toBe(true);
    } finally {
      clearMockRegistry();
      delete process.env.M365_ACCOUNT_HOME;
    }
  });
});

describe("--account", () => {
  const withRegistry = () => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work", hasSession: true }]);
  };

  it("writes the chosen account's record, not the active one's", async () => {
    withRegistry();
    const workStore = new PersonalStore(join(accountsDir(), "work", "session.json"));
    const defaultStore = new PersonalStore(join(accountsDir(), "default", "session.json"));
    const before = defaultStore.read();
    const out: string[] = [];
    const code = await runPersonalCli(["login", "--account", "work"], {
      capture: vi.fn(async () => capture),
      probe: vi.fn(async () => ({ model: "claude-sonnet", ok: true, at: 1 })),
      out: (l) => out.push(l), now: () => 1,
    });
    expect(code).toBe(0);
    expect(workStore.read()?.scenario).toBe("OfficeWebPaidConsumerCopilot");
    expect(defaultStore.read()).toEqual(before); // the active (default) account untouched
    expect(out.join("\n")).toContain('account: work');
    clearMockRegistry();
    delete process.env.M365_ACCOUNT_HOME;
  });

  it("status reads the chosen account", async () => {
    withRegistry();
    activateAccount("work");
    const out: string[] = [];
    await runPersonalCli(["status", "--account", "work", "--json"], { out: (l) => out.push(l), now: () => 1 });
    const parsed = JSON.parse(out.join("\n"));
    expect(parsed.scenario).toBeTruthy();
    expect(parsed.account).toBe("work");
    clearMockRegistry();
    delete process.env.M365_ACCOUNT_HOME;
  });

  it("--account with no value does not swallow the next flag", async () => {
    withRegistry();
    const out: string[] = [];
    const code = await runPersonalCli(["status", "--account", "--json"], { out: (l) => out.push(l), now: () => 1 });
    expect(code).toBe(2);
    expect(out.join("\n")).toContain("--account needs an account id");
    clearMockRegistry();
    delete process.env.M365_ACCOUNT_HOME;
  });

  it("an unknown account is a usage error, before anything is read or written", async () => {
    withRegistry();
    const captureFn = vi.fn(async () => capture);
    const out: string[] = [];
    expect(await runPersonalCli(["login", "--account", "nope"], { capture: captureFn, out: (l) => out.push(l) })).toBe(2);
    expect(captureFn).not.toHaveBeenCalled();
    expect(out.join("\n")).toContain("nope");
    clearMockRegistry();
    delete process.env.M365_ACCOUNT_HOME;
  });

  it("--account without a value is a usage error", async () => {
    withRegistry();
    const out: string[] = [];
    expect(await runPersonalCli(["status", "--account"], { out: (l) => out.push(l) })).toBe(2);
    clearMockRegistry();
    delete process.env.M365_ACCOUNT_HOME;
  });
});