import { capturePersonalSession, PersonalCaptureError } from "./personal-capture.js";
import { attachChatPage } from "./personal-harvest-cdp.js";
import { harvestPersonalToken, isHarvestError } from "./personal-harvest.js";
import type { PersonalCapture, PersonalRecord, PersonalRecordState } from "./personal-session.js";
import { PersonalStore } from "./personal-session.js";
import { resolveAccount, type AccountId } from "../../account-registry.js";
import { createLogger } from "../../io/log.js";
import {
  MINT_HOUR_MS,
  PERSONAL_MINT_RATE_LIMIT_ENV,
  personalMintRateLimits,
  type PersonalMintRateLimitOverride,
} from "./personal-mint-limits.js";

const log = createLogger("personal-auth");

export const PERSONAL_AUTH_ERROR_CODES = [
  "session_dead",
  "mint_rate_limited",
  "capture_failed",
  "chrome_unreachable",
  "no_chat_tab",
  "mint_not_rotated",
] as const;

export type PersonalAuthErrorCode = (typeof PERSONAL_AUTH_ERROR_CODES)[number];

export class PersonalAuthError extends Error {
  constructor(
    readonly code: PersonalAuthErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "PersonalAuthError";
  }
}

export type MintStrategy = "harvest" | "capture";

export interface PersonalAuthDeps {
  store: PersonalStore;
  accountId?: AccountId | null;
  record?: PersonalRecord | null;
  capture: (headed: boolean) => Promise<PersonalCapture>;
  harvest?: (prev: PersonalRecord) => Promise<{ capture: PersonalCapture; rotated: boolean } | null>;
  rateLimits?: PersonalMintRateLimitOverride;
  now: () => number;
}

export function personalLoginCommand(): string {
  return "pnpm m365-personal login";
}

function sessionDeadError(): PersonalAuthError {
  return new PersonalAuthError(
    "session_dead",
    `M365 personal session expired — run: ${personalLoginCommand()}`,
  );
}

function captureFailedError(reason: string): PersonalAuthError {
  if (isUnreachable(reason)) {
    return new PersonalAuthError(
      "chrome_unreachable",
      `could not reach Chrome on ${probeCdpTarget()} to re-mint the M365 personal session — start it on the capture profile: ${chromeCommandLine()}`,
    );
  }
  return new PersonalAuthError(
    "capture_failed",
    `could not re-mint the M365 personal session — run: ${personalLoginCommand()}`,
  );
}

export function probeCdpTarget(): string {
  return resolveAccount().cdpUrl;
}

export function chromeCommandLine(): string {
  const account = resolveAccount();
  return `chrome --remote-debugging-port=${new URL(account.cdpUrl).port || "9222"} --user-data-dir="${account.profileDir}" https://m365.cloud.microsoft/chat`;
}

function isUnreachable(reason: string): boolean {
  return /ECONNREFUSED|ECONNRESET|ETIMEDOUT|ENOTFOUND|connect over CDP|Target closed|browser has been closed/i.test(
    reason,
  );
}

function mintRateLimitedError(kind: "interval" | "hour"): PersonalAuthError {
  const what =
    kind === "interval"
      ? "the previous mint was less than the minimum mint interval ago"
      : "the hourly mint budget is exhausted";
  const limits = personalMintRateLimits();
  return new PersonalAuthError(
    "mint_rate_limited",
    `refusing to mint a fresh personal token: ${what} (the record is stale but not dead). Use ${PERSONAL_MINT_RATE_LIMIT_ENV.minIntervalMs} / ${PERSONAL_MINT_RATE_LIMIT_ENV.maxPerHour} to tune (currently ${limits.minIntervalMs} / ${limits.maxPerHour}), or ${personalLoginCommand()} to mint by hand.`,
  );
}

function defaultDeps(): PersonalAuthDeps {
  const account = resolveAccount();
  return {
    accountId: account.id,
    store: new PersonalStore(account.sessionFile),
    capture: (headed) => capturePersonalSession({ headed, profileDir: account.profileDir }),
    harvest: async (prev) => {
      if (mintStrategy() !== "harvest") return null;
      const result = await harvestPersonalToken({
        attach: attachChatPage,
        previousToken: prev.accessToken,
        now: () => Date.now(),
      });
      return { capture: result.capture, rotated: result.rotated };
    },
    now: () => Date.now(),
  };
}

function mintStrategy(): MintStrategy {
  return process.env.M365_PERSONAL_MINT_STRATEGY === "capture" ? "capture" : "harvest";
}

const inflight = new Map<string, Promise<string>>();

function stateKey(accountId: AccountId | null | undefined): string {
  return accountId ?? "builtin";
}

export async function getPersonalToken(deps?: Partial<PersonalAuthDeps>): Promise<string> {
  const d: PersonalAuthDeps = { ...defaultDeps(), ...deps };
  if (d.record && d.store.isFresh(d.record, d.now())) return d.record.accessToken;
  const state = inspectQuietly(d.store);
  const record = state.ok ? state.record : null;
  if (record && d.store.isFresh(record, d.now())) return record.accessToken;

  if (!state.ok && state.problem !== "absent") {
    log.error(`the stored personal session is unusable (${state.problem}: ${state.detail}); re-minting over it`);
  }

  const key = stateKey(d.accountId);
  let pending = inflight.get(key);
  if (!pending) {
    pending = mint(d, record).finally(() => {
      inflight.delete(key);
    });
    inflight.set(key, pending);
  }
  return pending;
}

function inspectQuietly(store: PersonalAuthDeps["store"]): PersonalRecordState {
  if (typeof store.inspect === "function") {
    try {
      return store.inspect();
    } catch (err) {
      return {
        ok: false,
        problem: "unreadable",
        detail: err instanceof Error ? err.message : String(err),
      };
    }
  }
  const record = store.read();
  return record
    ? { ok: true, record }
    : { ok: false, problem: "absent", detail: "no personal session file" };
}

async function mint(d: PersonalAuthDeps, prev: PersonalRecord | null): Promise<string> {
  const current = d.store.read();
  if (current && current !== prev && d.store.isFresh(current, d.now())) return current.accessToken;

  const now = d.now();
  const { minIntervalMs, maxPerHour } = personalMintRateLimits(process.env, d.rateLimits);
  if (prev) {
    if (now - prev.lastMintAt < minIntervalMs) {
      log.debug(`mint refused: only ${Math.round((now - prev.lastMintAt) / 1000)}s since the last mint`);
      throw mintRateLimitedError("interval");
    }
    const mintsLastHour = now - prev.lastMintAt < MINT_HOUR_MS ? prev.mintCount : 0;
    if (mintsLastHour >= maxPerHour) {
      log.debug(`mint refused: ${mintsLastHour} mints in the last hour (max ${maxPerHour})`);
      throw mintRateLimitedError("hour");
    }
  }

  const { capture, spentTurn } = await obtainToken(d, prev);

  const mintedAt = d.now();
  const record: PersonalRecord = {
    ...capture,
    emailHint: prev?.emailHint ?? null,
    probe: prev?.probe ?? null,
    mintCount: (prev ? (now - prev.lastMintAt < MINT_HOUR_MS ? prev.mintCount : 0) : 0) + (spentTurn ? 1 : 0),
    lastMintAt: spentTurn ? mintedAt : (prev?.lastMintAt ?? mintedAt),
  };
  d.store.write(record);
  log.debug(`minted, expires in ${Math.round((record.expiresAt - mintedAt) / 60_000)} min`);
  return record.accessToken;
}

async function obtainToken(
  d: PersonalAuthDeps,
  prev: PersonalRecord | null,
): Promise<{ capture: PersonalCapture; spentTurn: boolean }> {
  if (d.harvest) {
    try {
      const result = await d.harvest(prev ?? NO_PREVIOUS_RECORD);
      if (result) {
        if (!result.rotated) {
          log.debug("harvest: the page reused its current token (no rotation)");
        }
        return { capture: result.capture, spentTurn: true };
      }
    } catch (err) {
      const harvested = harvestErrorFor(err);
      if (harvested) throw harvested;
      log.debug(`harvest failed with an unexpected error: ${describeCaptureError(err)}; falling back to a capture`);
    }
  }
  return { capture: await launchCapture(d), spentTurn: false };
}

const NO_PREVIOUS_RECORD: PersonalRecord = {
  accessToken: "",
  chathubPath: "",
  scenario: "",
  variants: "",
  licenseType: "",
  agentHost: "",
  product: null,
  source: "officeweb",
  optionsSets: [],
  allowedMessageTypes: [],
  expiresAt: 0,
  capturedAt: 0,
  emailHint: null,
  probe: null,
  mintCount: 0,
  lastMintAt: 0,
};

async function launchCapture(d: PersonalAuthDeps): Promise<PersonalCapture> {
  log.debug("mint attempt (headless)");
  try {
    return await d.capture(false);
  } catch (err) {
    if (err instanceof PersonalCaptureError && err.reason === "not_chathub") {
      log.debug(`mint failed: the profile is dead (${err.message})`);
      throw sessionDeadError();
    }
    const headReason = describeCaptureError(err);
    log.debug(`mint failed headlessly: ${headReason}; retrying headed`);
    try {
      return await d.capture(true);
    } catch (retryErr) {
      const headedReason = describeCaptureError(retryErr);
      log.debug(`mint failed headed too: ${headedReason}`);
      throw captureFailedError(isUnreachable(headReason) && isUnreachable(headedReason) ? headedReason : "");
    }
  }
}

function harvestErrorFor(err: unknown): PersonalAuthError | null {
  if (!isHarvestError(err)) return null;
  switch (err.reason) {
    case "chrome_unreachable":
      return new PersonalAuthError(
        "chrome_unreachable",
        `could not reach Chrome on ${probeCdpTarget()} to re-mint the M365 personal session — start it on the capture profile: ${chromeCommandLine()}`,
      );
    case "no_chat_tab":
      return new PersonalAuthError(
        "no_chat_tab",
        `Chrome on ${probeCdpTarget()} has no m365.cloud.microsoft/chat tab open — open https://m365.cloud.microsoft/chat, or run: ${personalLoginCommand()}`,
      );
    case "session_dead":
      return sessionDeadError();
    case "no_chat_frame":
      return new PersonalAuthError(
        "mint_not_rotated",
        `a message was sent but the page opened no chat socket, so no token could be read — retry, or run: ${personalLoginCommand()}`,
      );
    case "capture_failed":
      return new PersonalAuthError(
        "capture_failed",
        `could not read a session from the attached Chrome (${err.message}) — run: ${personalLoginCommand()}`,
      );
  }
}

function describeCaptureError(err: unknown): string {
  if (err instanceof PersonalCaptureError) return `${err.reason}: ${err.message}`;
  if (err instanceof Error) return err.message;
  return String(err);
}

const cachedCredentials = new Map<string, Promise<void>>();

export function ensurePersonalCredentials(
  accountId?: AccountId | null,
  deps?: Partial<PersonalAuthDeps> & { getToken?: () => Promise<string> },
): Promise<void> {
  const key = stateKey(accountId);
  const cached = cachedCredentials.get(key);
  if (cached) return cached;
  const token = deps?.getToken
    ? deps.getToken()
    : getPersonalToken({
        store: new PersonalStore(resolveAccount({ id: accountId ?? null }).sessionFile),
        ...deps,
        accountId: accountId ?? deps?.accountId ?? null,
      });
  const pending = token.then(
    () => undefined,
    (err) => {
      cachedCredentials.delete(key);
      throw err;
    },
  );
  cachedCredentials.set(key, pending);
  return pending;
}

export function resetPersonalAuthState(accountId?: AccountId | null): void {
  if (accountId === undefined) {
    inflight.clear();
    cachedCredentials.clear();
    return;
  }
  const key = stateKey(accountId);
  inflight.delete(key);
  cachedCredentials.delete(key);
}
