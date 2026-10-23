import { chmodSync, copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, isAbsolute, join, relative, resolve } from "node:path";
import { PersonalStore, stripBom } from "./personal-session.js";
import { createLogger } from "./log.js";

const log = createLogger("accounts");

/** A slug, not a label: `[a-z0-9][a-z0-9-]{0,31}`. Validated before anything
 *  touches the filesystem, because it becomes a directory name. */
export type AccountId = string;

/** One account: where its record lives, and which Chrome profile can re-mint it.
 *  Both paths are stored absolute under the accounts container so the operator
 *  can see them; the registry never derives a path from the id. */
export interface AccountEntry {
  id: AccountId;
  label: string;
  cdpPort: number;
  /** Absolute path of the account's session record. */
  sessionFile: string;
  /** Absolute path of the account's Chrome profile directory. */
  profileDir: string;
  /** ISO 8601. */
  createdAt: string;
}

export interface AccountRegistry {
  version: 1;
  activeId: AccountId;
  accounts: AccountEntry[];
}

/** The pair of things an account is: where its record lives, and which Chrome can re-mint it. */
export type AccountPaths = Pick<AccountEntry, "id" | "sessionFile" | "profileDir" | "cdpPort">;

export const ACCOUNT_ID_RE: RegExp = /^[a-z0-9][a-z0-9-]{0,31}$/;
/** Today's record file name, kept by the migration so no file moves. */
export const ACCOUNT_SESSION_FILE = "personal-session.json";
/** Today's Chrome profile name, likewise. */
export const ACCOUNT_PROFILE_DIR = "personal-profile";
/** The first port an account is offered; a second account takes the next free port upward. */
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

/** Where accounts added later live: `<config>/accounts`. */
export function accountsDir(): string {
  return join(configDir(), "accounts");
}

/** The registry file's default location: `<config>/accounts.json`. */
export function accountsRoot(): string {
  return join(configDir(), "accounts.json");
}

/** Reads the registry file's raw text. `null` means "no registry yet"; every
 *  other failure is rethrown so the registry layer can name it. */
export interface AccountRegistrySource {
  readonly file: string;
  /** What `file` would be with no `M365_ACCOUNTS_FILE` and no explicit file:
   *  `<config>/accounts.json`. Lets a caller say where the registry lives. */
  readonly defaultFile: string;
  read(): string | null;
}

/** M365_ACCOUNTS_FILE, else `<config>/accounts.json`. Not cached: the caller
 *  reads through it on every use, which is what lets a switch take effect on
 *  the next request. */
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

/** The first free port at or above the default: a port every entry already
 *  names is taken (a second account must not share a debugging port), and the
 *  first free one upward is handed out. */
function nextFreePort(accounts: AccountEntry[]): number {
  const taken = new Set(accounts.map((a) => a.cdpPort));
  let port = ACCOUNT_CDP_PORT;
  while (taken.has(port)) {
    if (port >= MAX_CDP_PORT) throw new Error("no free CDP port left to allocate");
    port += 1;
  }
  return port;
}

/** The last port the allocator will consider. */
const MAX_CDP_PORT = 65_535;

/** Build the entry for a new account under `accounts/<id>/`. Throws when the id
 *  is already in the registry or the account's session file already exists.
 *  A bare directory is NOT identity: `addAccount` creates the pair, and a
 *  failed or half-finished attempt must not block the retry that fixes it. */
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

/** A record-shaped file, enough for `existsSync`/`PersonalStore` to see a
 *  session. Never used by production code. */
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

/** Write the real files the flags ask for under `<config>/accounts/<id>/`, then
 *  the registry that points at them, into `source.file`. */
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
    mkdirSync(dirname(entry.sessionFile), { recursive: true });
    if (mock.hasSession) writeFileSync(entry.sessionFile, JSON.stringify(MOCK_SESSION, null, 2));
    if (mock.profileExists) mkdirSync(entry.profileDir, { recursive: true });
    entries.push(entry);
  }
  const active = activeId ?? accounts[0]?.id ?? "default";
  const registry: AccountRegistry = { version: 1, activeId: active, accounts: entries };
  writeFileSync(source.file, JSON.stringify(registry, null, 2) + "\n");
}

/** A registry file in a fresh temp dir, at exactly this path, with the real
 *  files the flags ask for. */
export function createMockRegistry(accounts: MockAccount[] = [], activeId?: string): AccountRegistrySource {
  return createMockRegistryAt(join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json"), accounts, activeId);
}

/** `createMockRegistry` at a call-chosen path — the variant `useMockRegistry`
 *  needs so the file it writes is the file `M365_ACCOUNTS_FILE` names. */
export function createMockRegistryAt(
  file: string,
  accounts: MockAccount[] = [],
  activeId?: string,
): AccountRegistrySource {
  const source = accountRegistrySource(file);
  writeMockAccounts(accounts, activeId, source);
  return source;
}

/** Point the process at a new temporary registry and fill it in. */
export function useMockRegistry(accounts: MockAccount[] = []): AccountRegistrySource {
  const file = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
  process.env.M365_ACCOUNTS_FILE = file;
  return createMockRegistryAt(file, accounts);
}

/** Drop the `M365_ACCOUNTS_FILE` override, so the next call resolves the real
 *  config dir (or is re-pointed by the test). */
export function clearMockRegistry(): void {
  delete process.env.M365_ACCOUNTS_FILE;
}

/** `~` and `~/x` (or `~\x`) expand to the home directory; anything else
 *  relative resolves against the config dir, so a hand-written
 *  `"personal-session.json"` in the registry means the same file the old
 *  layout used. */
export function expandHome(p: string): string {
  let out = p;
  if (p === "~") out = homedir();
  else if (p.startsWith("~/") || p.startsWith("~\\")) out = join(homedir(), p.slice(2));
  return isAbsolute(out) ? out : resolve(configDir(), out);
}

export function isSafeAccountId(id: string): boolean {
  return ACCOUNT_ID_RE.test(id);
}

/** The outcome of one registry read. `problem` is `"absent"` when there is no
 *  registry yet, otherwise a short machine-readable code (`invalid_json`,
 *  `wrong_shape`) the caller turns into a message. */
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

/** Read and normalise the registry. Paths come back absolute (`~` expanded);
 *  every failure is reported, never thrown — the read paths must answer even
 *  when the operator's file is broken. */
export function readRegistry(src?: AccountRegistrySource): RegistryRead {
  const source = src ?? accountRegistrySource();
  let raw: string | null;
  try {
    raw = source.read();
  } catch (err) {
    // A path that exists but cannot be read (a directory, a lock, permissions)
    // is not "absent": the fix is different, and Task 3 says so.
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

/** Write the registry: temp file (0o600) then rename, so a reader never sees a
 *  half-written file. A rename that cannot cross a volume (`EXDEV`) falls back
 *  to copy + unlink; either way no `.tmp-*` file is left behind. */
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

/** The `M365_PERSONAL_*` variables `migrateRegistry` cares about. */
const PERSONAL_OVERRIDES = ["M365_PERSONAL_SESSION_FILE", "M365_PERSONAL_PROFILE", "M365_PERSONAL_CDP_URL"] as const;

function hasPersonalOverride(env: Record<string, string | undefined>): boolean {
  return PERSONAL_OVERRIDES.some((k) => Boolean(env[k]));
}

/** A registry with no readable account to fall back on. Kept as the best
 *  effort for a malformed existing registry: the caller reports it, and a
 *  write path refuses rather than overwriting the operator's file. */
function emptyRegistry(): AccountRegistry {
  return { version: 1, activeId: "default", accounts: [] };
}

/** Create the registry on first use, pointing `default` at today's files.
 *
 *  Migration is a pointer, never a move: with no override the entry names
 *  `personal-session.json` / `personal-profile` / port 9222 exactly where they
 *  already are, so a Chrome that is running during the upgrade keeps working.
 *  When the operator does set an override, the pair is placed under
 *  `accounts/default/` so a later unsetting of the variable cannot move it. */
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

/** What is wrong with a registry the operator may have hand-edited. `null`
 *  when nothing is. */
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

/** Where the accounts come from right now, and what is wrong if anything is.
 *  This is what `/healthz` and `/accounts` report (§7.9), and it never throws:
 *  the proxy has to keep answering with a broken registry. */
export interface RegistryStatus {
  /** The registry file's absolute path, as resolved for this process (§7.9). */
  path: string;
  /** `registry` = accounts.json; `builtin` = no usable registry (M365_PERSONAL_* / today's defaults). */
  repository: "registry" | "builtin";
  activeId: string | null;
  count: number;
  /** Why the registry could not be used, and what to do about it. */
  problem?: { code: RegistryProblemCode; detail: string; hint: string };
}

/** Report the registry without ever throwing. A read failure (a directory where
 *  the file should be, a lock, permissions) becomes `unreadable`; anything else
 *  `readRegistry` would report is passed through. */
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

/** Every account id, in registry order; `[]` when there is no usable registry. */
export function accountIds(src?: AccountRegistrySource): string[] {
  return readRegistry(src).registry?.accounts.map((a) => a.id) ?? [];
}

/** The Microsoft identity a record names: `chathubPath` is `oid@tid`, so the
 *  segment before the `@` is the account's `oid` — the key quota and throttling
 *  actually use. `null` when the path carries none. */
export function chathubOid(chathubPath: string): string | null {
  const at = chathubPath.indexOf("@");
  if (at <= 0) return null;
  const oid = chathubPath.slice(0, at).toLowerCase();
  return oid.length > 0 ? oid : null;
}

/** One account as the list route and the UI need it. */
export interface AccountView {
  id: AccountId;
  label: string;
  active: boolean;
  cdpPort: number;
  sessionFile: string;
  profileDir: string;
  profileExists: boolean;
  /** The Microsoft identity, when a usable record is stored. */
  oid: string | null;
  /** Ids of other accounts whose stored record names the same `oid`. */
  oidCollisionWith: AccountId[];
  /** Set when the two share one daily turn budget; `null` otherwise. */
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

/** The list projection: every entry plus its profile presence and identity.
 *  Never throws — an unreadable record is simply "no oid". */
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

/** Why an account operation was refused. The route layer maps these to status
 *  codes (Task 11); each carries what to do about it. */
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

/** Read the registry for a write: a missing registry is created (migration),
 *  a broken one refuses rather than being overwritten. */
function readOrCreate(src?: AccountRegistrySource): { source: AccountRegistrySource; registry: AccountRegistry } {
  const source = src ?? accountRegistrySource();
  const read = readRegistry(source);
  if (read.ok) return { source, registry: read.registry };
  if (read.problem === "absent") {
    // Creating the registry is this write's own job; `migrateRegistry` would
    // only refuse to touch what is already there (and has none — the read above
    // just said so).
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

/** The exact command that opens an account's own Chrome: its profile and its
 *  debugging port (mirrors `chromeCommandLine()` in `personal-auth.ts`). */
export function chromeCommandFor(account: Pick<AccountEntry, "profileDir" | "cdpPort">): string {
  return `chrome --remote-debugging-port=${account.cdpPort} --user-data-dir="${account.profileDir}" https://m365.cloud.microsoft/chat`;
}

export function loginCommandFor(id: AccountId): string {
  return `pnpm m365-personal login --account ${id}`;
}

/** Add an account: validate, allocate, create its directory pair, record it.
 *  A duplicate Microsoft account is a warning, never a block (§7.4): the two
 *  would share one turn budget, and the operator is the one who knows why. */
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
  // A brand-new account has no record, so an oid can only collide once it has
  // logged in; `accountViews()` carries the warning from then on.
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

/** Change what the UI shows. The id and both paths are immutable. */
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

/** Point new requests and new conversations at this account. A running session
 *  keeps the account it started on (§7.3). */
export function activateAccount(id: AccountId, src?: AccountRegistrySource): AccountRegistry {
  const { source, registry } = readOrCreate(src);
  findEntry(registry, id);
  const next: AccountRegistry = { ...registry, activeId: id };
  writeRegistry(next, source);
  return next;
}

/** Remove an account by moving its files aside — never deleting them. Refuses
 *  the active account (switch first), the last account (nothing would serve),
 *  and anything the caller found busy (the route checks the pool). */
export function removeAccount(
  id: AccountId,
  src?: AccountRegistrySource,
  now: Date = new Date(),
): { trashed: string } {
  const { source, registry } = readOrCreate(src);
  const entry = findEntry(registry, id);
  // Last first: with one account left, "it is the active one" is true but
  // useless — the operator needs to know nothing would serve afterwards.
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
    // Today's layout: the files are not under the account's own directory, so
    // move each half that exists into one trash directory.
    mkdirSync(trashDir, { recursive: true, mode: 0o700 });
    if (existsSync(entry.sessionFile)) renameSync(entry.sessionFile, join(trashDir, "session.json"));
    if (existsSync(entry.profileDir)) renameSync(entry.profileDir, join(trashDir, "profile"));
  }
  writeRegistry({ ...registry, accounts: registry.accounts.filter((a) => a.id !== id) }, source);
  return { trashed: trashDir };
}

/** Whether `p` is `dir` or lives under it. */
function isInside(p: string, dir: string): boolean {
  const rel = relative(dir, p);
  return rel !== "" && !rel.startsWith("..") && !isAbsolute(rel);
}
