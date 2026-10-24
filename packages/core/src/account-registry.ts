import { chmodSync, copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, isAbsolute, join, relative, resolve } from "node:path";
import { PersonalStore, stripBom } from "./adapters/personal/personal-session.js";
import { createLogger } from "./io/log.js";

const log = createLogger("accounts");

/** A slug, not a label: `[a-z0-9][a-z0-9-]{0,31}`. Validated before anything
 *  touches the filesystem, because it becomes a directory name. */
export type AccountId = string;

export interface AccountEntry {
  id: AccountId;
  label: string;
  cdpPort: number;
  sessionFile: string;
  profileDir: string;
  createdAt: string;
}

export interface AccountRegistry {
  version: 1;
  activeId: AccountId;
  accounts: AccountEntry[];
}

export type AccountPaths = Pick<AccountEntry, "id" | "sessionFile" | "profileDir" | "cdpPort">;

export const ACCOUNT_ID_RE: RegExp = /^[a-z0-9][a-z0-9-]{0,31}$/;
export const ACCOUNT_SESSION_FILE = "personal-session.json";
export const ACCOUNT_PROFILE_DIR = "personal-profile";
export const ACCOUNT_CDP_PORT = 9222;

/** The per-user config dir, computed per call so a different `HOME` is seen
 *  (`turn-budget.ts` does the same). `M365_ACCOUNT_HOME` points the whole
 *  account layout somewhere else — for tests, and for a second set of accounts
 *  on one machine (§7.9) — and is inert when unset. */
export function configDir(): string {
  const base = process.env.M365_ACCOUNT_HOME;
  if (base) return resolve(expandHome(base));
  return join(homedir(), ".config", "opencode-m365");
}

export function accountsDir(): string {
  return join(configDir(), "accounts");
}

export function accountsRoot(): string {
  return join(configDir(), "accounts.json");
}

export interface AccountRegistrySource {
  readonly file: string;
  readonly defaultFile: string;
  read(): string | null;
}

export function accountRegistrySource(file?: string): AccountRegistrySource {
  const defaultFile = accountsRoot();
  const path = file ?? process.env.M365_ACCOUNTS_FILE ?? defaultFile;
  return {
    file: path,
    defaultFile,
    read(): string | null {
      try {
        return readFileSync(path, "utf8");
      } catch (err) {
        if ((err as NodeJS.ErrnoException)?.code === "ENOENT") return null;
        throw err;
      }
    },
  };
}

function nextFreePort(accounts: AccountEntry[]): number {
  const taken = new Set(accounts.map((a) => a.cdpPort));
  let port = ACCOUNT_CDP_PORT;
  while (taken.has(port)) {
    if (port >= MAX_CDP_PORT) throw new Error("no free CDP port left to allocate");
    port += 1;
  }
  return port;
}

const MAX_CDP_PORT = 65_535;

export function allocateAccount(
  reg: AccountRegistry,
  id: AccountId,
  label: string,
  now: Date,
): AccountEntry {
  if (!ACCOUNT_ID_RE.test(id)) {
    throw new Error(`invalid account id "${id}": ids match ${ACCOUNT_ID_RE.source}`);
  }
  const taken = reg.accounts.some((a) => a.id === id) || existsSync(join(accountsDir(), id, "session.json"));
  if (taken) throw new Error(`account "${id}" already exists`);
  return {
    id,
    label,
    cdpPort: nextFreePort(reg.accounts),
    sessionFile: join(accountsDir(), id, "session.json"),
    profileDir: join(accountsDir(), id, "profile"),
    createdAt: now.toISOString(),
  };
}

export interface MockAccount {
  id: string;
  hasSession?: boolean;
  profileExists?: boolean;
  cdpPort?: number;
}

const MOCK_SESSION = {
  accessToken: "mock-token",
  chathubPath: "mock@tenant",
  scenario: "mock",
  variants: "mock",
  licenseType: "Free",
  agentHost: "mock",
  product: null,
  source: "officeweb",
  optionsSets: [],
  allowedMessageTypes: ["Chat"],
  expiresAt: 0,
  capturedAt: 0,
  emailHint: null,
  probe: null,
  mintCount: 0,
  lastMintAt: 0,
};

const REAL_HOME_DIR = homedir();
const REAL_CONFIG_DIR = join(homedir(), ".config", "opencode-m365");

export function isRealDataDir(p: string | undefined): boolean {
  if (!p) return true;
  const real = REAL_CONFIG_DIR.replace(/\\/g, "/").toLowerCase().replace(/\/+$/, "");
  const path = resolve(p).replace(/\\/g, "/").toLowerCase().replace(/\/+$/, "");
  return path === real || path.startsWith(`${real}/`) || path === REAL_HOME_DIR.replace(/\\/g, "/").toLowerCase();
}

function writeMockAccounts(
  accounts: MockAccount[],
  activeId: string | undefined,
  source: AccountRegistrySource,
): void {
  mkdirSync(dirname(source.file), { recursive: true });
  const entries: AccountEntry[] = [];
  for (const mock of accounts) {
    const entry: AccountEntry = {
      id: mock.id,
      label: mock.id,
      cdpPort: mock.cdpPort ?? ACCOUNT_CDP_PORT + entries.length,
      sessionFile: join(accountsDir(), mock.id, "session.json"),
      profileDir: join(accountsDir(), mock.id, "profile"),
      createdAt: new Date().toISOString(),
    };
    if (isRealDataDir(entry.sessionFile) || isRealDataDir(entry.profileDir) || isRealDataDir(source.file)) {
      throw new Error(
        `mock registry would write into the real data dir (${REAL_CONFIG_DIR}); point M365_ACCOUNT_HOME (or use useMockRegistry) at a temp dir first`,
      );
    }
    mkdirSync(dirname(entry.sessionFile), { recursive: true });
    if (mock.hasSession) writeFileSync(entry.sessionFile, JSON.stringify(MOCK_SESSION, null, 2));
    if (mock.profileExists) mkdirSync(entry.profileDir, { recursive: true });
    entries.push(entry);
  }
  const active = activeId ?? accounts[0]?.id ?? "default";
  const registry: AccountRegistry = { version: 1, activeId: active, accounts: entries };
  writeFileSync(source.file, JSON.stringify(registry, null, 2) + "\n");
}

export function createMockRegistry(accounts: MockAccount[] = [], activeId?: string): AccountRegistrySource {
  return createMockRegistryAt(join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json"), accounts, activeId);
}

export function createMockRegistryAt(
  file: string,
  accounts: MockAccount[] = [],
  activeId?: string,
): AccountRegistrySource {
  const source = accountRegistrySource(file);
  writeMockAccounts(accounts, activeId, source);
  return source;
}

// `M365_ACCOUNT_HOME` points at the temp dir because `writeMockAccounts`
// creates `<config>/accounts/<id>/`, so a mock built against the real config
// dir would create directories inside the operator's live signed-in account.
export function useMockRegistry(accounts: MockAccount[] = []): AccountRegistrySource {
  const dir = mkdtempSync(join(tmpdir(), "ar-"));
  process.env.M365_ACCOUNT_HOME = dir;
  const file = join(dir, "accounts.json");
  process.env.M365_ACCOUNTS_FILE = file;
  return createMockRegistryAt(file, accounts);
}

export function clearMockRegistry(): void {
  delete process.env.M365_ACCOUNTS_FILE;
  delete process.env.M365_ACCOUNT_HOME;
}

export function expandHome(p: string): string {
  let out = p;
  if (p === "~") out = homedir();
  else if (p.startsWith("~/") || p.startsWith("~\\")) out = join(homedir(), p.slice(2));
  return isAbsolute(out) ? out : resolve(configDir(), out);
}

export function isSafeAccountId(id: string): boolean {
  return ACCOUNT_ID_RE.test(id);
}

export type RegistryRead =
  | { ok: true; registry: AccountRegistry }
  | { ok: false; problem: "absent" | string; registry: null };

function isEntry(value: unknown): value is AccountEntry {
  if (typeof value !== "object" || value === null) return false;
  const e = value as Record<string, unknown>;
  return (
    typeof e.id === "string" &&
    typeof e.label === "string" &&
    typeof e.cdpPort === "number" &&
    Number.isFinite(e.cdpPort) &&
    typeof e.sessionFile === "string" &&
    typeof e.profileDir === "string" &&
    typeof e.createdAt === "string"
  );
}

export function readRegistry(src?: AccountRegistrySource): RegistryRead {
  const source = src ?? accountRegistrySource();
  let raw: string | null;
  try {
    raw = source.read();
  } catch (err) {
    return { ok: false, problem: "unreadable", registry: null };
  }
  if (raw === null) return { ok: false, problem: "absent", registry: null };
  let parsed: unknown;
  try {
    parsed = JSON.parse(stripBom(raw));
  } catch {
    return { ok: false, problem: "invalid_json", registry: null };
  }
  if (typeof parsed !== "object" || parsed === null) return { ok: false, problem: "wrong_shape", registry: null };
  const r = parsed as Record<string, unknown>;
  if (r.version !== 1 || typeof r.activeId !== "string" || !Array.isArray(r.accounts)) {
    return { ok: false, problem: "wrong_shape", registry: null };
  }
  const accounts: AccountEntry[] = [];
  for (const entry of r.accounts) {
    if (!isEntry(entry)) return { ok: false, problem: "wrong_shape", registry: null };
    accounts.push({
      ...entry,
      sessionFile: expandHome(entry.sessionFile),
      profileDir: expandHome(entry.profileDir),
    });
  }
  return { ok: true, registry: { version: 1, activeId: r.activeId, accounts } };
}

export function writeRegistry(registry: AccountRegistry, src?: AccountRegistrySource): void {
  const source = src ?? accountRegistrySource();
  const path = source.file;
  mkdirSync(dirname(path), { recursive: true, mode: 0o700 });
  const tmp = `${path}.tmp-${process.pid}`;
  const raw = JSON.stringify(registry, null, 2) + "\n";
  writeFileSync(tmp, raw, { mode: 0o600 });
  try {
    renameSync(tmp, path);
  } catch {
    try {
      copyFileSync(tmp, path);
    } finally {
      try {
        unlinkSync(tmp);
      } catch (err) {
        log.debug("removing the registry temp file failed:", err);
      }
    }
  }
  if (process.platform !== "win32") {
    try {
      chmodSync(path, 0o600);
    } catch (err) {
      log.debug("chmod 0o600 on the registry failed:", err);
    }
  }
}

export interface MigrationResult {
  created: boolean;
  registry: AccountRegistry;
}

const PERSONAL_OVERRIDES = ["M365_PERSONAL_SESSION_FILE", "M365_PERSONAL_PROFILE", "M365_PERSONAL_CDP_URL"] as const;

function hasPersonalOverride(env: Record<string, string | undefined>): boolean {
  return PERSONAL_OVERRIDES.some((k) => Boolean(env[k]));
}

function emptyRegistry(): AccountRegistry {
  return { version: 1, activeId: "default", accounts: [] };
}

// Migration is a pointer, never a move: with no override the entry names
// `personal-session.json` / `personal-profile` / port 9222 exactly where they
// already are, so a Chrome that is running during the upgrade keeps working.
export function migrateRegistry(src?: AccountRegistrySource, now: Date = new Date()): MigrationResult {
  const source = src ?? accountRegistrySource();
  const read = readRegistry(source);
  if (read.ok) return { created: false, registry: read.registry };
  if (read.problem !== "absent") return { created: false, registry: emptyRegistry() };

  const override = hasPersonalOverride(process.env);
  const entry: AccountEntry = {
    id: "default",
    label: "Personal (free)",
    cdpPort: ACCOUNT_CDP_PORT,
    sessionFile: override
      ? join(accountsDir(), "default", "session.json")
      : join(configDir(), ACCOUNT_SESSION_FILE),
    profileDir: override
      ? join(accountsDir(), "default", "profile")
      : join(configDir(), ACCOUNT_PROFILE_DIR),
    createdAt: now.toISOString(),
  };
  const registry: AccountRegistry = { version: 1, activeId: "default", accounts: [entry] };
  writeRegistry(registry, source);
  return { created: true, registry };
}

export interface AccountProblem {
  code: "duplicate_id" | "unknown_active" | "no_accounts";
  detail: string;
}

export function registryProblem(registry: AccountRegistry): AccountProblem | null {
  const ids = registry.accounts.map((a) => a.id);
  if (registry.accounts.length === 0) {
    return { code: "no_accounts", detail: "the registry lists no accounts" };
  }
  const dupes = ids.filter((id, i) => ids.indexOf(id) !== i);
  if (dupes.length > 0) {
    return {
      code: "duplicate_id",
      detail: `more than one account has the id${dupes.length > 1 ? "s" : ""}: ${[...new Set(dupes)].join(", ")}`,
    };
  }
  if (!ids.includes(registry.activeId)) {
    return {
      code: "unknown_active",
      detail: `activeId "${registry.activeId}" names no account (known: ${ids.join(", ")})`,
    };
  }
  return null;
}

export type RegistryProblemCode = "absent" | "invalid_json" | "wrong_shape" | "unreadable" | string;

export interface RegistryStatus {
  path: string;
  repository: "registry" | "builtin";
  activeId: string | null;
  count: number;
  problem?: { code: RegistryProblemCode; detail: string; hint: string };
}

export function registryStatus(src?: AccountRegistrySource): RegistryStatus {
  const source = src ?? accountRegistrySource();
  let read = readRegistry(source);
  if (!read.ok && read.registry === null && read.problem !== "absent") {
    try {
      source.read();
    } catch (err) {
      read = { ok: false, problem: "unreadable", registry: null };
      return builtinStatus(source.file, {
        code: "unreadable",
        detail: String(err),
        hint: invalidHint(source.file),
      });
    }
  }
  if (!read.ok) {
    return builtinStatus(source.file, {
      code: read.problem,
      detail:
        read.problem === "absent"
          ? `no account registry at ${source.file}`
          : `${source.file} is not a usable registry (${read.problem})`,
      hint:
        read.problem === "absent"
          ? "no account registry yet — the registry is created on first use of personal mode"
          : invalidHint(source.file),
    });
  }
  const problem = registryProblem(read.registry);
  if (problem) {
    return builtinStatus(source.file, {
      code: problem.code,
      detail: problem.detail,
      hint: invalidHint(source.file),
    });
  }
  return {
    path: source.file,
    repository: "registry",
    activeId: read.registry.activeId,
    count: read.registry.accounts.length,
  };
}

function invalidHint(path: string): string {
  return `fix or delete ${path}; the proxy is using the builtin layout until it is fixed`;
}

function builtinStatus(path: string, problem: NonNullable<RegistryStatus["problem"]>): RegistryStatus {
  log.debug(`registry unusable (${problem.code}): ${problem.detail}`);
  return { path, repository: "builtin", activeId: null, count: 0, problem };
}

export function accountIds(src?: AccountRegistrySource): string[] {
  return readRegistry(src).registry?.accounts.map((a) => a.id) ?? [];
}

export function chathubOid(chathubPath: string): string | null {
  const at = chathubPath.indexOf("@");
  if (at <= 0) return null;
  const oid = chathubPath.slice(0, at).toLowerCase();
  return oid.length > 0 ? oid : null;
}

export interface AccountView {
  id: AccountId;
  label: string;
  active: boolean;
  cdpPort: number;
  sessionFile: string;
  profileDir: string;
  profileExists: boolean;
  oid: string | null;
  oidCollisionWith: AccountId[];
  warning: string | null;
}

function profileExistsAt(dir: string): boolean {
  try {
    return statSync(dir).isDirectory();
  } catch {
    return false;
  }
}

function oidOf(entry: AccountEntry): string | null {
  const state = new PersonalStore(entry.sessionFile).inspect();
  return state.ok ? chathubOid(state.record.chathubPath) : null;
}

export function accountViews(src?: AccountRegistrySource): AccountView[] {
  const reg = readRegistry(src).registry;
  if (!reg) return [];
  const oids = reg.accounts.map(oidOf);
  return reg.accounts.map((entry, i) => {
    const oid = oids[i] ?? null;
    const oidCollisionWith =
      oid === null ? [] : reg.accounts.filter((other, j) => j !== i && oids[j] === oid).map((other) => other.id);
    return {
      id: entry.id,
      label: entry.label,
      active: entry.id === reg.activeId,
      cdpPort: entry.cdpPort,
      sessionFile: entry.sessionFile,
      profileDir: entry.profileDir,
      profileExists: profileExistsAt(entry.profileDir),
      oid,
      oidCollisionWith,
      warning:
        oidCollisionWith.length > 0
          ? `shares Microsoft account ${oid} with ${oidCollisionWith.join(", ")} — they share one daily turn budget`
          : null,
    };
  });
}

export class AccountError extends Error {
  constructor(
    readonly code: "invalid_id" | "not_found" | "already_exists" | "active" | "last_account" | "registry_invalid",
    message: string,
    readonly hint: string,
  ) {
    super(message);
    this.name = "AccountError";
  }
}

export interface NewAccount {
  entry: AccountEntry;
  chromeCommand: string;
  loginCommand: string;
  warning: string | null;
}

function readOrCreate(src?: AccountRegistrySource): { source: AccountRegistrySource; registry: AccountRegistry } {
  const source = src ?? accountRegistrySource();
  const read = readRegistry(source);
  if (read.ok) return { source, registry: read.registry };
  if (read.problem === "absent") {
    const entry: AccountEntry = {
      id: "default",
      label: "Personal (free)",
      cdpPort: ACCOUNT_CDP_PORT,
      sessionFile: join(configDir(), ACCOUNT_SESSION_FILE),
      profileDir: join(configDir(), ACCOUNT_PROFILE_DIR),
      createdAt: new Date().toISOString(),
    };
    return { source, registry: { version: 1, activeId: "default", accounts: [entry] } };
  }
  throw new AccountError(
    "registry_invalid",
    `the account registry at ${source.file} cannot be used (${read.problem})`,
    invalidHint(source.file),
  );
}

function findEntry(registry: AccountRegistry, id: AccountId): AccountEntry {
  const entry = registry.accounts.find((a) => a.id === id);
  if (!entry) {
    throw new AccountError(
      "not_found",
      `account "${id}" does not exist`,
      `known accounts: ${registry.accounts.map((a) => a.id).join(", ") || "none"}`,
    );
  }
  return entry;
}

export function chromeCommandFor(account: Pick<AccountEntry, "profileDir" | "cdpPort">): string {
  return `chrome --remote-debugging-port=${account.cdpPort} --user-data-dir="${account.profileDir}" https://m365.cloud.microsoft/chat`;
}

export function loginCommandFor(id: AccountId): string {
  return `pnpm m365-personal login --account ${id}`;
}

export function addAccount(
  id: AccountId,
  label: string,
  src?: AccountRegistrySource,
  now: Date = new Date(),
): NewAccount {
  if (!isSafeAccountId(id)) {
    throw new AccountError(
      "invalid_id",
      `invalid account id "${id}"`,
      "ids match [a-z0-9][a-z0-9-]{0,31}: lowercase letters, digits and dashes, starting with a letter or digit",
    );
  }
  const { source, registry } = readOrCreate(src);
  if (registry.accounts.some((a) => a.id === id)) {
    throw new AccountError("already_exists", `account "${id}" already exists`, "pick another id, or rename the existing one");
  }
  let entry: AccountEntry;
  try {
    entry = allocateAccount(registry, id, label, now);
  } catch (err) {
    throw new AccountError("already_exists", (err as Error).message, "pick another id, or remove the leftover directory");
  }
  mkdirSync(accountsDir(), { recursive: true, mode: 0o700 });
  mkdirSync(entry.profileDir, { recursive: true, mode: 0o700 });
  writeRegistry({ ...registry, accounts: [...registry.accounts, entry] }, source);
  const oid = oidOf(entry);
  const collision = oid === null ? [] : registry.accounts.filter((a) => oidOf(a) === oid).map((a) => a.id);
  return {
    entry,
    chromeCommand: chromeCommandFor(entry),
    loginCommand: loginCommandFor(id),
    warning:
      collision.length > 0
        ? `shares Microsoft account ${oid} with ${collision.join(", ")} — they share one daily turn budget`
        : null,
  };
}

export function renameAccount(id: AccountId, label: string, src?: AccountRegistrySource): AccountEntry {
  const { source, registry } = readOrCreate(src);
  const entry = findEntry(registry, id);
  const renamed = { ...entry, label };
  writeRegistry(
    { ...registry, accounts: registry.accounts.map((a) => (a.id === id ? renamed : a)) },
    source,
  );
  return renamed;
}

export function activateAccount(id: AccountId, src?: AccountRegistrySource): AccountRegistry {
  const { source, registry } = readOrCreate(src);
  findEntry(registry, id);
  const next: AccountRegistry = { ...registry, activeId: id };
  writeRegistry(next, source);
  return next;
}

export function removeAccount(
  id: AccountId,
  src?: AccountRegistrySource,
  now: Date = new Date(),
): { trashed: string } {
  const { source, registry } = readOrCreate(src);
  const entry = findEntry(registry, id);
  if (registry.accounts.length <= 1) {
    throw new AccountError(
      "last_account",
      `account "${id}" is the last one`,
      "an account is the last one; add another first",
    );
  }
  if (registry.activeId === id) {
    throw new AccountError(
      "active",
      `account "${id}" is the active account`,
      "activate another account first",
    );
  }
  const trashDir = join(accountsDir(), ".trash", `${id}-${now.getTime()}`);
  mkdirSync(dirname(trashDir), { recursive: true, mode: 0o700 });
  const owned = join(accountsDir(), id);
  if (isInside(entry.sessionFile, accountsDir()) && isInside(entry.profileDir, accountsDir())) {
    renameSync(owned, trashDir);
  } else {
    mkdirSync(trashDir, { recursive: true, mode: 0o700 });
    if (existsSync(entry.sessionFile)) renameSync(entry.sessionFile, join(trashDir, "session.json"));
    if (existsSync(entry.profileDir)) renameSync(entry.profileDir, join(trashDir, "profile"));
  }
  writeRegistry({ ...registry, accounts: registry.accounts.filter((a) => a.id !== id) }, source);
  return { trashed: trashDir };
}

function isInside(p: string, dir: string): boolean {
  const rel = relative(dir, p);
  return rel !== "" && !rel.startsWith("..") && !isAbsolute(rel);
}

export interface ResolvedAccount {
  id: AccountId | null;
  sessionFile: string;
  profileDir: string;
  cdpUrl: string;
  source: "cli" | "env" | "registry" | "builtin";
}

export interface ResolveAccountOptions {
  id?: AccountId | null;
  env?: Record<string, string | undefined>;
  onWarn?: (line: string) => void;
}

const PERSONAL_SETTINGS = [
  { env: "M365_PERSONAL_SESSION_FILE", what: "session file", fallback: () => join(configDir(), "personal-session.json") },
  { env: "M365_PERSONAL_PROFILE", what: "browser profile", fallback: () => join(configDir(), "personal-profile") },
  { env: "M365_PERSONAL_CDP_URL", what: "CDP target", fallback: () => `http://127.0.0.1:${ACCOUNT_CDP_PORT}` },
] as const;

const warned = new Set<string>();

function warnOnce(key: string, line: string, onWarn?: (line: string) => void): boolean {
  if (warned.has(key)) return false;
  warned.add(key);
  if (onWarn) onWarn(line);
  else console.error(`[accounts] ${line}`);
  return true;
}

export function resolveAccount(opts: ResolveAccountOptions = {}): ResolvedAccount {
  const env = opts.env ?? process.env;
  const builtin: Omit<ResolvedAccount, "id" | "source"> = {
    sessionFile: env.M365_PERSONAL_SESSION_FILE || PERSONAL_SETTINGS[0].fallback(),
    profileDir: env.M365_PERSONAL_PROFILE || PERSONAL_SETTINGS[1].fallback(),
    cdpUrl: env.M365_PERSONAL_CDP_URL || PERSONAL_SETTINGS[2].fallback(),
  };

  const source = accountRegistrySource();
  const read = readRegistry(source);
  const registry = read.registry;
  const overrides = PERSONAL_SETTINGS.filter((s) => Boolean(env[s.env]));

  if (opts.id) {
    if (!read.ok && read.problem !== "absent") {
      throw new AccountError(
        "registry_invalid",
        `cannot resolve account "${opts.id}": the account registry is unusable (${read.problem})`,
        invalidHint(source.file),
      );
    }
    if (opts.id !== null && !isSafeAccountId(opts.id)) {
      throw new AccountError(
        "invalid_id",
        `invalid account id "${opts.id}"`,
        `ids match ${ACCOUNT_ID_RE.source}: lowercase letters, digits and dashes, starting with a letter or digit`,
      );
    }
    const entry = registry?.accounts.find((a) => a.id === opts.id);
    if (!entry) {
      const known = registry?.accounts.map((a) => a.id) ?? [];
      throw new AccountError(
        "not_found",
        `account "${opts.id}" does not exist`,
        known.length > 0
          ? `known accounts: ${known.join(", ")}`
          : `no accounts yet - add one with the proxy's /accounts/ui (or POST /accounts)`,
      );
    }
    return resolvedFrom(entry, "cli");
  }

  if (overrides.length > 0) {
    let reported = false;
    for (const s of overrides) {
      reported =
        warnOnce(
          s.env,
          registry
            ? `${s.env} overrides the registry; account switching will not change the ${s.what}`
            : `${s.env} overrides the default ${s.what}`,
          opts.onWarn,
        ) || reported;
    }
    if (overrides.length < PERSONAL_SETTINGS.length) {
      const line = `only ${overrides.map((s) => s.env).join(", ")} ${overrides.length > 1 ? "are" : "is"} set - a partial override splits the record/profile/CDP pair; set all three or none`;
      if (!warnOnce("partial", line, opts.onWarn) && opts.onWarn && reported) opts.onWarn(line);
    }
    return { id: null, source: "env", ...builtin };
  }

  if (registry) {
    const entry = registry.accounts.find((a) => a.id === registry.activeId) ?? registry.accounts[0];
    if (entry) {
      return resolvedFrom(entry, "registry");
    }
  }

  try {
    mkdirSync(configDir(), { recursive: true });
  } catch (err) {
    log.debug("creating the config dir failed:", err);
  }
  return { id: null, source: "builtin", ...builtin };
}

function resolvedFrom(entry: AccountEntry, source: ResolvedAccount["source"]): ResolvedAccount {
  return {
    id: entry.id,
    sessionFile: entry.sessionFile,
    profileDir: entry.profileDir,
    cdpUrl: `http://127.0.0.1:${entry.cdpPort}`,
    source,
  };
}

export function accountKeyOf(id: AccountId | null | undefined): string | null {
  return id ?? null;
}
