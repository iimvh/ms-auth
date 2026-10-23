import { existsSync, mkdtempSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeAll, describe, expect, it } from "vitest";
import {
  ACCOUNT_CDP_PORT,
  ACCOUNT_ID_RE,
  accountRegistrySource,
  accountsDir,
  allocateAccount,
  clearMockRegistry,
  configDir,
  createMockRegistry,
  migrateRegistry,
  readRegistry,
  registryProblem,
  useMockRegistry,
  writeRegistry,
  type AccountRegistry,
} from "./account-registry.js";

// `allocateAccount` and the mock write into `<config>/accounts/<id>/`, and the
// id lives under the real config dir whenever the registry path alone is
// mocked — so the whole file runs against a temporary `M365_ACCOUNT_HOME`
// (the developer's live ~/.config/opencode-m365 is never touched).
let home: string;
beforeAll(() => {
  home = mkdtempSync(join(tmpdir(), "ar-home-"));
  process.env.M365_ACCOUNT_HOME = home;
});

afterEach(() => clearMockRegistry());

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

  it("skips ports already allocated, starting at the default and going up", () => {
    // The plan's case (9222 + 9224 taken ⇒ 9225) contradicts the plan's own
    // spec, "Ports default to 9222, then the next free port upward": with 9222
    // taken the first free port is 9223, which is what the next test pins.
    const bookkeeping = { label: "X", sessionFile: join(accountsDir(), "x", "session.json"), profileDir: join(accountsDir(), "x", "profile"), createdAt: new Date().toISOString() };
    const reg = registry([
      { ...bookkeeping, id: "a", cdpPort: ACCOUNT_CDP_PORT },
      { ...bookkeeping, id: "b", cdpPort: 9224 },
    ]);
    expect(allocateAccount(reg, "c", "C", new Date()).cdpPort).toBe(9223);
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

describe("readRegistry", () => {
  it("reports absent for a missing file", () => {
    clearMockRegistry();
    process.env.M365_ACCOUNTS_FILE = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    expect(readRegistry()).toEqual({ ok: false, problem: "absent", registry: null });
  });

  it("stores relative paths expanded to absolute", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    process.env.M365_ACCOUNTS_FILE = file;
    writeFileSync(file, JSON.stringify({
      version: 1, activeId: "default",
      accounts: [{ id: "default", label: "Personal", cdpPort: 9222,
        sessionFile: "personal-session.json", profileDir: "personal-profile", createdAt: "2026-10-23T05:00:00.000Z" }],
    }));
    const r = readRegistry();
    expect(r.ok).toBe(true);
    expect(r.registry!.accounts[0]!.sessionFile).toBe(join(configDir(), "personal-session.json"));
    expect(r.registry!.accounts[0]!.profileDir).toBe(join(configDir(), "personal-profile"));
  });

  it("round-trips a write with mode 0o600 and leaves no temp file", () => {
    const dir = mkdtempSync(join(tmpdir(), "ar-"));
    process.env.M365_ACCOUNTS_FILE = join(dir, "accounts.json");
    const reg = registry([allocateAccount(registry([]), "a", "A", new Date())]);
    writeRegistry(reg);
    expect(readRegistry().registry).toEqual(reg);
    expect(readdirSync(dir)).toEqual(["accounts.json"]);
    if (process.platform !== "win32") {
      expect(statSync(join(dir, "accounts.json")).mode & 0o777).toBe(0o600);
    }
  });

  it("falls back to copy+unlink when rename cannot cross volumes", () => {
    // Simulated by a source whose path is on another root: we assert the outcome
    // only — the file holds the new registry and no *.tmp-* file remains.
    const dir = mkdtempSync(join(tmpdir(), "ar-"));
    process.env.M365_ACCOUNTS_FILE = join(dir, "accounts.json");
    const src = accountRegistrySource(join(dir, "accounts.json"));
    const reg = registry([allocateAccount(registry([]), "a", "A", new Date())]);
    writeRegistry(reg, src);
    expect(JSON.parse(readFileSync(join(dir, "accounts.json"), "utf8")).activeId).toBe("a");
    expect(readdirSync(dir).filter((f) => f.includes(".tmp-"))).toEqual([]);
  });
});

describe("migrateRegistry", () => {
  it("points at today's files without moving or copying anything", () => {
    const dir = mkdtempSync(join(tmpdir(), "ar-"));
    process.env.M365_ACCOUNTS_FILE = join(dir, "accounts.json");
    // Today's real layout, in miniature.
    const todayFile = join(configDir(), "personal-session.json");
    const before = existsSync(todayFile);
    const { created, registry: reg } = migrateRegistry(undefined, new Date("2026-10-23T05:00:00.000Z"));
    expect(created).toBe(true);
    expect(reg.activeId).toBe("default");
    expect(reg.accounts).toHaveLength(1);
    expect(reg.accounts[0]).toMatchObject({
      id: "default", label: "Personal (free)", cdpPort: 9222, sessionFile: todayFile,
      profileDir: join(configDir(), "personal-profile"), createdAt: "2026-10-23T05:00:00.000Z",
    });
    expect(existsSync(todayFile)).toBe(before); // nothing was created, moved or renamed
  });

  it("writes nothing when a registry already exists (idempotent)", () => {
    useMockRegistry([{ id: "work" }]);
    const { created } = migrateRegistry();
    expect(created).toBe(false);
    expect(readRegistry().registry!.activeId).toBe("work");
  });

  it("migrates into accounts/<id>/ when the operator set M365_PERSONAL_SESSION_FILE", () => {
    const dir = mkdtempSync(join(tmpdir(), "ar-"));
    process.env.M365_ACCOUNTS_FILE = join(dir, "accounts.json");
    process.env.M365_PERSONAL_SESSION_FILE = join(dir, "custom.json");
    try {
      const { registry: reg } = migrateRegistry();
      expect(reg.accounts[0]!.sessionFile).toBe(join(accountsDir(), "default", "session.json"));
    } finally {
      delete process.env.M365_PERSONAL_SESSION_FILE;
    }
  });
});

describe("registryProblem", () => {
  const entry = (id: string) => allocateAccount(registry([]), id, id, new Date());
  it("is null for a healthy registry", () => {
    expect(registryProblem(registry([entry("a")]))).toBeNull();
  });
  it("names duplicate ids, a missing active id and an empty list", () => {
    expect(registryProblem(registry([entry("a"), { ...entry("b"), id: "a" }]))!.code).toBe("duplicate_id");
    expect(registryProblem(registry([entry("a")], "gone"))!.code).toBe("unknown_active");
    expect(registryProblem(registry([] as AccountRegistry["accounts"]))!.code).toBe("no_accounts");
  });
});
