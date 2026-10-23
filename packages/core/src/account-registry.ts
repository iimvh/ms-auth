import { existsSync, mkdirSync, mkdtempSync, readFileSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, join } from "node:path";

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
 *  (`turn-budget.ts` does the same). */
export function configDir(): string {
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

/** A registry file in a fresh temp dir, with the real files the flags ask for. */
export function createMockRegistry(accounts: MockAccount[] = [], activeId?: string): AccountRegistrySource {
  const source = accountRegistrySource(join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json"));
  writeMockAccounts(accounts, activeId, source);
  return source;
}

/** Point the process at a new temporary registry and fill it in. */
export function useMockRegistry(accounts: MockAccount[] = []): AccountRegistrySource {
  process.env.M365_ACCOUNTS_FILE = join(mkdtempSync(join(tmpdir(), "ar-")), "accounts.json");
  return createMockRegistry(accounts);
}

/** Drop the `M365_ACCOUNTS_FILE` override, so the next call resolves the real
 *  config dir (or is re-pointed by the test). */
export function clearMockRegistry(): void {
  delete process.env.M365_ACCOUNTS_FILE;
}
