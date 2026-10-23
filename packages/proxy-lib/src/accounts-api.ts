import {
  AccountError,
  accountViews,
  activateAccount,
  addAccount,
  chathubOid,
  isPersonalMode,
  loginCommandFor,
  PersonalStore,
  personalHealth,
  probeCdpReachable,
  registryStatus,
  removeAccount,
  renameAccount,
  resolveAccount,
  type AccountView,
  type PersonalHealth,
  type PersonalHealthDeps,
  type RegistryStatus,
} from "@m365-copilot/core";

/**
 * The account-management API (spec §7.5), as pure functions over injected deps
 * so the routes are thin and this is testable without a Nitro server.
 *
 * Every route answers 404 in enterprise (MSAL) mode: there are no personal
 * accounts to manage there, and a 404 is the least surprising answer for a
 * feature that does not exist in that deployment.
 */

export interface AccountHealthView extends AccountView {
  /** The account's own credential state (its record + its Chrome), or null
   *  when the probe could not run. */
  credential: PersonalHealth | null;
  /** Set when the account shares an identity with another one (one budget). */
  warning: string | null;
  /** A conversation opened on this account is still in the pool (so it cannot
   *  be deleted yet, and the dashboard shows why). */
  accountIdLive: boolean;
}

export interface AccountsDeps {
  pool: { liveAccountIds(): string[]; forgetAccount(id: string): number };
  now?: () => number;
  /** How long a CDP probe is reused, so a dashboard's refresh loop does not
   *  open a socket per account per second. */
  healthCacheMs?: number;
  /** The CDP probe, injected so tests never touch a real port. */
  probeCdp?: (url: string) => Promise<boolean>;
}

export interface RouteResult {
  status: number;
  body: unknown;
}

export interface AccountsListBody {
  registry: RegistryStatus;
  accounts: AccountHealthView[];
}

const DEFAULT_HEALTH_CACHE_MS = 10_000;

// --- The 10 s probe cache ---

interface CacheEntry { at: number; url: string; value: PersonalHealth }
const healthCache = new Map<string, CacheEntry>();

/** Only for tests: forget every cached probe. */
export function resetAccountsHealthCache(): void {
  healthCache.clear();
}

async function healthOf(accountId: string, deps: AccountsDeps): Promise<PersonalHealth | null> {
  const now = (deps.now ?? Date.now)();
  const cacheMs = deps.healthCacheMs ?? DEFAULT_HEALTH_CACHE_MS;
  // The account's own paths, whatever the env says about the active one.
  const resolved = resolveAccount({ id: accountId });
  const cached = healthCache.get(accountId);
  if (cached && cached.url === resolved.cdpUrl && now - cached.at < cacheMs) return cached.value;
  const d: Partial<PersonalHealthDeps> = { now: () => now };
  // personalHealth keys its record and CDP probe off resolveAccount(), i.e. the
  // ACTIVE account; a listed account may not be that one, so give it its own.
  if (deps.probeCdp) d.probeCdp = deps.probeCdp;
  if (accountId !== resolveAccount().id) {
    d.store = new PersonalStore(resolved.sessionFile);
    if (!deps.probeCdp) d.probeCdp = () => probeCdpReachable(resolved.cdpUrl);
  }
  try {
    const value = await personalHealth(d);
    healthCache.set(accountId, { at: now, url: resolved.cdpUrl, value });
    return value;
  } catch {
    return null;
  }
}

// --- Helpers ---

/** Whether that account is serving a conversation right now. */
export function isAccountLive(deps: AccountsDeps, id: string): boolean {
  return deps.pool.liveAccountIds().includes(id);
}

function errorBody(code: string, message: string, hint?: string, extra?: Record<string, unknown>): RouteResult["body"] {
  return { error: { code, message, ...(hint ? { hint } : {}), ...extra } };
}

/** `true` while any conversation of that account is still in the pool. */
function labelProblem(label: unknown): string | null {
  if (typeof label !== "string") return "label must be a string";
  if (!label.trim()) return "label must not be empty";
  if (label.length > 120) return "label must be at most 120 characters";
  return null;
}

/**
 * Run a route body, mapping `AccountError` to its HTTP status and anything
 * unexpected to a 500 (rather than letting it escape as a Nitro HTML page).
 */
export async function toRouteResult(fn: () => RouteResult | Promise<RouteResult>): Promise<RouteResult> {
  try {
    return await fn();
  } catch (err) {
    if (err instanceof AccountError) {
      const status =
        err.code === "invalid_id" ? 400
          : err.code === "not_found" ? 404
            : 409; // already_exists | active | last_account | registry_invalid
      return { status, body: errorBody(err.code, err.message, err.hint) };
    }
    const message = err instanceof Error ? err.message : String(err);
    return { status: 500, body: errorBody("internal", message) };
  }
}

// --- Routes ---

export async function listAccounts(deps: AccountsDeps): Promise<AccountsListBody> {
  const views = accountViews();
  const accounts: AccountHealthView[] = [];
  for (const view of views) {
    accounts.push({
      ...view,
      credential: await healthOf(view.id, deps),
      warning: view.warning,
      accountIdLive: isAccountLive(deps, view.id),
    });
  }
  return { registry: registryStatus(), accounts };
}

export function accountsListResponse(deps: AccountsDeps): Promise<RouteResult> {
  if (!isPersonalMode()) return Promise.resolve({ status: 404, body: errorBody("not_found", "account management requires personal auth mode") });
  return toRouteResult(async () => ({ status: 200, body: await listAccounts(deps) }));
}

export async function accountsAddResponse(
  input: { id: unknown; label: unknown },
  deps: AccountsDeps,
): Promise<RouteResult> {
  if (!isPersonalMode()) return { status: 404, body: errorBody("not_found", "account management requires personal auth mode") };
  return toRouteResult(() => {
    const bad = labelProblem(input.label);
    if (bad) return { status: 400, body: errorBody("invalid_label", bad) };
    const id = typeof input.id === "string" ? input.id.trim() : "";
    if (!id) return { status: 400, body: errorBody("invalid_id", "id is required") };
    // addAccount validates the id shape and the duplicate; readRegistry's own
    // failures come back as AccountError("registry_invalid").
    const created = addAccount(id, (input.label as string).trim());
    return {
      status: 201,
      body: {
        entry: created.entry,
        chromeCommand: created.chromeCommand,
        loginCommand: created.loginCommand,
        warning: created.warning,
      },
    };
  });
}

export function accountsRenameResponse(id: string, input: { label: unknown }, deps: AccountsDeps): RouteResult {
  if (!isPersonalMode()) return { status: 404, body: errorBody("not_found", "account management requires personal auth mode") };
  void deps;
  const bad = labelProblem(input.label);
  if (bad) return { status: 400, body: errorBody("invalid_label", bad) };
  try {
    const entry = renameAccount(id, (input.label as string).trim());
    return { status: 200, body: { entry } };
  } catch (err) {
    return synchronously(err, () => ({ status: 200, body: {} }));
  }
}

export function accountsActivateResponse(id: string, deps: AccountsDeps): RouteResult {
  if (!isPersonalMode()) return { status: 404, body: errorBody("not_found", "account management requires personal auth mode") };
  void deps;
  try {
    const registry = activateAccount(id);
    return { status: 200, body: { activeId: registry.activeId } };
  } catch (err) {
    return synchronously(err, () => ({ status: 200, body: {} }));
  }
}

export function accountsDeleteResponse(id: string, deps: AccountsDeps): { result: RouteResult; forgotten: number } {
  if (!isPersonalMode()) {
    return { result: { status: 404, body: errorBody("not_found", "account management requires personal auth mode") }, forgotten: 0 };
  }
  // Before touching the registry: a conversation opened on this account cannot
  // be re-keyed onto another identity, and removing the entry would strand it.
  // A conversation stays in the pool until it has been idle past
  // M365_SESSION_IDLE_MS (60 min by default), so the honest instruction is to
  // switch the client to another account (or stop using that thread) and let
  // the idle sweep let it go — the page cannot finish it.
  if (isAccountLive(deps, id)) {
    return {
      result: {
        status: 409,
        body: errorBody(
          "account_busy",
          `account "${id}" is serving a live conversation`,
          "that thread keeps this account until it has been idle past the session idle limit (M365_SESSION_IDLE_MS, 60 min by default) — switch your client to another account or restart the proxy, then delete again",
          { conversationAccount: id },
        ),
      },
      forgotten: 0,
    };
  }
  try {
    removeAccount(id);
    // Its pool state is now stale (a re-created id must not inherit it).
    const forgotten = deps.pool.forgetAccount(id);
    return { result: { status: 200, body: { removed: id, forgotten } }, forgotten };
  } catch (err) {
    return { result: synchronously(err, () => ({ status: 200, body: {} })), forgotten: 0 };
  }
}

/** The synchronous twin of toRouteResult, for the routes that do no I/O. */
function synchronously(err: unknown, _ok: () => RouteResult): RouteResult {
  if (err instanceof AccountError) {
    const status = err.code === "invalid_id" ? 400 : err.code === "not_found" ? 404 : 409;
    return { status, body: errorBody(err.code, err.message, err.hint) };
  }
  return { status: 500, body: errorBody("internal", err instanceof Error ? err.message : String(err)) };
}

/** `chathubOid` re-exported for the UI, which shows the identity per account. */
export { chathubOid, loginCommandFor };
