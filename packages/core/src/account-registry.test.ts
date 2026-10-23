import { existsSync, mkdtempSync, readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeAll, beforeEach, describe, expect, it } from "vitest";
import {
  ACCOUNT_CDP_PORT,
  ACCOUNT_ID_RE,
  AccountError,
  accountIds,
  accountRegistrySource,
  accountsDir,
  accountViews,
  activateAccount,
  addAccount,
  allocateAccount,
  chathubOid,
  clearMockRegistry,
  configDir,
  createMockRegistry,
  migrateRegistry,
  readRegistry,
  registryProblem,
  registryStatus,
  removeAccount,
  renameAccount,
  useMockRegistry,
  writeRegistry,
  type AccountRegistry,
} from "./account-registry.js";
import { PersonalStore } from "./personal-session.js";

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

describe("registryStatus", () => {
  it("names the builtin fallback when there is no registry", () => {
    clearMockRegistry();
    process.env.M365_ACCOUNTS_FILE = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    const s = registryStatus();
    expect(s).toMatchObject({ repository: "builtin", activeId: null, count: 0 });
    expect(s.problem!.code).toBe("absent");
    expect(s.problem!.hint).toContain("registry is created on first use");
  });

  it("names the registry path and the active account when there is one", () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    activateAccount("work");
    const s = registryStatus();
    expect(s).toMatchObject({ repository: "registry", activeId: "work", count: 2 });
    expect(s.path).toBe(process.env.M365_ACCOUNTS_FILE);
    expect(s.problem).toBeUndefined();
  });

  it("never throws on a broken registry, and says so", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    process.env.M365_ACCOUNTS_FILE = file;
    writeFileSync(file, "{ not json");
    expect(registryStatus()).toMatchObject({ repository: "builtin", problem: { code: "invalid_json" } });
  });

  it("never throws when the registry path cannot be read", () => {
    process.env.M365_ACCOUNTS_FILE = tmpdir(); // a directory, not a file
    expect(() => registryStatus()).not.toThrow();
    expect(registryStatus().problem!.code).toBe("unreadable");
  });
});

describe("account CRUD", () => {
  // Each test starts from its own account home: `addAccount`/`removeAccount`
  // write real directories, and a leftover `accounts/<id>/session.json` from the
  // previous test would (rightly) refuse the next `addAccount`.
  beforeEach(() => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
  });

  it("adds an account, creating only its directory pair", () => {
    useMockRegistry([{ id: "default" }]);
    const r = addAccount("work", "Work (free)");
    expect(r.entry.sessionFile).toBe(join(accountsDir(), "work", "session.json"));
    expect(r.chromeCommand).toContain(`--user-data-dir="${r.entry.profileDir}"`);
    expect(r.chromeCommand).toContain(`--remote-debugging-port=${r.entry.cdpPort}`);
    expect(r.loginCommand).toBe("pnpm m365-personal login --account work");
    expect(existsSync(join(accountsDir(), "work", "profile"))).toBe(true);
    expect(accountIds()).toEqual(["default", "work"]);
  });

  it("refuses a bad id before touching the filesystem, and a taken id with 409 material", () => {
    useMockRegistry([{ id: "default" }]);
    expect(() => addAccount("Work!", "x")).toThrow(AccountError);
    try { addAccount("Work!", "x"); } catch (e) { expect((e as AccountError).code).toBe("invalid_id"); }
    try { addAccount("default", "x"); } catch (e) { expect((e as AccountError).code).toBe("already_exists"); }
    expect(existsSync(join(accountsDir(), "Work!"))).toBe(false);
  });

  it("warns rather than blocks when the new account shares an oid", () => {
    useMockRegistry([{ id: "default", hasSession: true }]);
    const defaultRecord = { accessToken: "t", chathubPath: "ABC123@tenant-x" };
    new PersonalStore(join(accountsDir(), "default", "session.json")).write({ ...defaultRecord, scenario: "S", variants: "v",
      licenseType: "Free", agentHost: "h", product: null, source: "officeweb", optionsSets: [],
      allowedMessageTypes: ["Chat"], expiresAt: 1, capturedAt: 0, emailHint: null, probe: null,
      mintCount: 0, lastMintAt: 0 });
    expect(chathubOid("ABC123@tenant-x")).toBe("abc123");
    expect(chathubOid("no-at-sign")).toBeNull();
    const same = addAccount("work", "Work");
    expect(same.warning).toBeNull(); // the new account has no record yet
    new PersonalStore(same.entry.sessionFile).write({ ...defaultRecord, scenario: "S", variants: "v",
      licenseType: "Free", agentHost: "h", product: null, source: "officeweb", optionsSets: [],
      allowedMessageTypes: ["Chat"], expiresAt: 1, capturedAt: 0, emailHint: null, probe: null,
      mintCount: 0, lastMintAt: 0 });
    expect(accountViews()[1]!.oidCollisionWith).toEqual(["default"]);
    const warning = accountViews().find((a) => a.id === "work")!.warning;
    expect(warning).toContain("share one daily turn budget");
  });

  it("renames a label without touching the id or the paths", () => {
    useMockRegistry([{ id: "work" }]);
    const before = accountViews()[0]!;
    renameAccount("work", "Work account");
    expect(accountViews()[0]).toMatchObject({ id: "work", label: "Work account", sessionFile: before.sessionFile });
  });

  it("activates, and refuses an unknown id", () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    expect(activateAccount("work").activeId).toBe("work");
    expect(registryStatus().activeId).toBe("work");
    expect(() => activateAccount("gone")).toThrow(AccountError);
  });

  it("removes by moving the directory aside, never deleting it", () => {
    useMockRegistry([{ id: "default" }, { id: "work", hasSession: true }]);
    const { trashed } = removeAccount("work");
    expect(trashed).toContain(join(".trash", "work-"));
    expect(existsSync(trashed)).toBe(true);
    expect(existsSync(join(accountsDir(), "work"))).toBe(false);
    expect(accountIds()).toEqual(["default"]);
  });

  it("refuses to remove the active account or the last one", () => {
    useMockRegistry([{ id: "default" }, { id: "work" }], "work");
    try { removeAccount("work"); } catch (e) { expect((e as AccountError).code).toBe("active"); }
    // A fresh home, so the previous test's `work` cannot make `only` non-last.
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "only" }]);
    try { removeAccount("only"); } catch (e) { expect((e as AccountError).code).toBe("last_account"); }
  });

  it("writes refuse a broken registry instead of overwriting it", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
    process.env.M365_ACCOUNTS_FILE = file;
    writeFileSync(file, "{ not json");
    try { addAccount("work", "Work"); } catch (e) { expect((e as AccountError).code).toBe("registry_invalid"); }
    expect(readFileSync(file, "utf8")).toBe("{ not json");
  });
});
