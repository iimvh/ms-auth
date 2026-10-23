import { existsSync, mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import {
  ACCOUNT_CDP_PORT,
  ACCOUNT_ID_RE,
  accountRegistrySource,
  accountsDir,
  allocateAccount,
  clearMockRegistry,
  configDir,
  createMockRegistry,
  useMockRegistry,
  type AccountRegistry,
} from "./account-registry.js";

afterEach(() => clearMockRegistry());

// `allocateAccount` (rightly) refuses a directory that already exists, and the
// id lives under the real config dir — so tests that allocate must not see each
// other's leftovers. Both dirs are test-only artefacts (plan Task 1 ruling).
afterEach(() => {
  for (const id of ["work", "a", "b", "c", "only"]) {
    rmSync(join(accountsDir(), id), { recursive: true, force: true });
  }
});

const registry = (accounts: AccountRegistry["accounts"], activeId = accounts[0]?.id ?? "default"): AccountRegistry =>
  ({ version: 1, activeId, accounts });

describe("account ids", () => {
  it("accepts lowercase slugs and rejects anything else", () => {
    expect(ACCOUNT_ID_RE.test("default")).toBe(true);
    expect(ACCOUNT_ID_RE.test("w0rk-2")).toBe(true);
    for (const bad of ["Work", "-work", "wo rk", "", "a".repeat(33), "../etc", "wörk"]) {
      expect(ACCOUNT_ID_RE.test(bad)).toBe(false);
    }
  });
});

describe("allocateAccount", () => {
  it("puts a new account under accounts/<id>/ at the default port", () => {
    const a = allocateAccount(registry([]), "work", "Work (free)", new Date("2026-10-23T06:00:00.000Z"));
    expect(a.sessionFile).toBe(join(accountsDir(), "work", "session.json"));
    expect(a.profileDir).toBe(join(accountsDir(), "work", "profile"));
    expect(a.cdpPort).toBe(ACCOUNT_CDP_PORT);
    expect(a.createdAt).toBe("2026-10-23T06:00:00.000Z");
  });

  it("gives the second account the next free port", () => {
    const first = allocateAccount(registry([]), "a", "A", new Date());
    const reg = registry([first]);
    expect(allocateAccount(reg, "b", "B", new Date()).cdpPort).toBe(first.cdpPort + 1);
  });

  it("skips ports already allocated, above the default too", () => {
    const base = { ...allocateAccount(registry([]), "a", "A", new Date()), cdpPort: ACCOUNT_CDP_PORT };
    const above = {
      ...base,
      id: "b",
      sessionFile: join(accountsDir(), "b", "session.json"),
      profileDir: join(accountsDir(), "b", "profile"),
      cdpPort: 9224,
    };
    expect(allocateAccount(registry([base, above]), "c", "C", new Date()).cdpPort).toBe(9225);
  });

  it("refuses an id that is already taken", () => {
    const first = allocateAccount(registry([]), "a", "A", new Date());
    expect(() => allocateAccount(registry([first]), "a", "B", new Date())).toThrow(/already/i);
  });
});

describe("the registry source", () => {
  it("is M365_ACCOUNTS_FILE when set, else the config dir", () => {
    clearMockRegistry();
    expect(accountRegistrySource().read()).toBeNull();
    expect(accountRegistrySource().defaultFile).toBe(join(configDir(), "accounts.json"));
    const src = useMockRegistry([]);
    expect(src.read()).not.toBeNull();
  });

  it("re-reads the file on every call, so a rewrite is seen at once", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    const src = accountRegistrySource(file);
    expect(src.read()).toBeNull();
    process.env.M365_ACCOUNTS_FILE = file;
    const mock = createMockRegistry([{ id: "default" }, { id: "work" }]);
    expect(JSON.parse(mock.read()!).accounts).toHaveLength(2);
  });

  it("a mock's entries point where their flags say, and are absolute", () => {
    const src = useMockRegistry([{ id: "work", hasSession: true, profileExists: false, cdpPort: 9230 }]);
    const reg = JSON.parse(src.read()!) as AccountRegistry;
    expect(reg.activeId).toBe("work");
    expect(reg.accounts[0]!.sessionFile).toBe(join(accountsDir(), "work", "session.json"));
    expect(reg.accounts[0]!.profileDir).toBe(join(accountsDir(), "work", "profile"));
    expect(reg.accounts[0]!.cdpPort).toBe(9230);
    expect(existsSync(join(accountsDir(), "work", "session.json"))).toBe(true);
    expect(existsSync(join(accountsDir(), "work", "profile"))).toBe(false);
  });
});
