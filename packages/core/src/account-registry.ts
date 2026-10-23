import { chmodSync, copyFileSync, existsSync, mkdirSync, mkdtempSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, isAbsolute, join, resolve } from "node:path";
import { stripBom } from "./personal-session.js";
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
 *  is already in the registry, or when the account's own session file — its
 *  identity on disk — already exists. A bare leftover directory is not
 *  identity: the caller creates the pair (Task 3's `addAccount`), and a failed
 *  attempt must not block the retry. */
export function allocateAccount(
  reg: AccountRegistry,
  id: AccountId,
  label: string,
  now: Date,
): AccountEntry {
  if (!ACCOUNT_ID_RE.test(id)) {
    throw new Error(`invalid account id "${id}": ids match ${ACCOUNT_ID_RE.source}`);
  }
  const taken = reg.accounts.some((a) => a.id === id) || existsSync(join(accountsDir(), id));
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
