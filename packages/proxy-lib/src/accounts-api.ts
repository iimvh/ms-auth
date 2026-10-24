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

export interface AccountHealthView extends AccountView {
  credential: PersonalHealth | null;
  warning: string | null;
  accountIdLive: boolean;
}

export interface AccountsDeps {
  pool: { liveAccountIds(): string[]; forgetAccount(id: string): number };
  now?: () => number;
  healthCacheMs?: number;
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

interface CacheEntry { at: number; url: string; value: PersonalHealth }
const healthCache = new Map<string, CacheEntry>();

export function resetAccountsHealthCache(): void {
  healthCache.clear();
}

async function healthOf(accountId: string, deps: AccountsDeps): Promise<PersonalHealth | null> {
  const now = (deps.now ?? Date.now)();
  const cacheMs = deps.healthCacheMs ?? DEFAULT_HEALTH_CACHE_MS;
  const resolved = resolveAccount({ id: accountId });
  const cached = healthCache.get(accountId);
  if (cached && cached.url === resolved.cdpUrl && now - cached.at < cacheMs) return cached.value;
  const d: Partial<PersonalHealthDeps> = { now: () => now };
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

export function isAccountLive(deps: AccountsDeps, id: string): boolean {
  return deps.pool.liveAccountIds().includes(id);
}

function errorBody(code: string, message: string, hint?: string, extra?: Record<string, unknown>): RouteResult["body"] {
  return { error: { code, message, ...(hint ? { hint } : {}), ...extra } };
}

function labelProblem(label: unknown): string | null {
  if (typeof label !== "string") return "label must be a string";
  if (!label.trim()) return "label must not be empty";
  if (label.length > 120) return "label must be at most 120 characters";
  return null;
}

export async function toRouteResult(fn: () => RouteResult | Promise<RouteResult>): Promise<RouteResult> {
  try {
    return await fn();
  } catch (err) {
    if (err instanceof AccountError) {
      const status =
        err.code === "invalid_id" ? 400
          : err.code === "not_found" ? 404
            : 409;
      return { status, body: errorBody(err.code, err.message, err.hint) };
    }
    const message = err instanceof Error ? err.message : String(err);
    return { status: 500, body: errorBody("internal", message) };
  }
}

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
    const forgotten = deps.pool.forgetAccount(id);
    return { result: { status: 200, body: { removed: id, forgotten } }, forgotten };
  } catch (err) {
    return { result: synchronously(err, () => ({ status: 200, body: {} })), forgotten: 0 };
  }
}

function synchronously(err: unknown, _ok: () => RouteResult): RouteResult {
  if (err instanceof AccountError) {
    const status = err.code === "invalid_id" ? 400 : err.code === "not_found" ? 404 : 409;
    return { status, body: errorBody(err.code, err.message, err.hint) };
  }
  return { status: 500, body: errorBody("internal", err instanceof Error ? err.message : String(err)) };
}

export { chathubOid, loginCommandFor };
