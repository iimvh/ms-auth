import { capturePersonalSession, PersonalCaptureError } from "./personal-capture.js";
import { attachChatPage } from "./personal-harvest-cdp.js";
import { harvestPersonalToken, isHarvestError } from "./personal-harvest.js";
import type { PersonalCapture, PersonalRecord, PersonalRecordState } from "./personal-session.js";
import { PersonalStore, personalProfileDir, personalSessionFile } from "./personal-session.js";
import { resolveAccount } from "./account-registry.js";
import { createLogger } from "./log.js";

const log = createLogger("personal-auth");

export type PersonalAuthErrorCode =
  | "session_dead"
  | "mint_rate_limited"
  | "capture_failed"
  | "chrome_unreachable"
  | "no_chat_tab"
  | "mint_not_rotated";

export class PersonalAuthError extends Error {
  constructor(
    readonly code: PersonalAuthErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "PersonalAuthError";
  }
}

/** `harvest` forces the page to open a socket and reads the token off it — the
 *  mechanism that works for a Chrome already running (F111, approach C of
 *  docs/personal-mint-design.md). `capture` launches a browser and is what
 *  `login` uses. */
export type MintStrategy = "harvest" | "capture";

export interface PersonalAuthDeps {
  store: PersonalStore;
  capture: (headed: boolean) => Promise<PersonalCapture>;
  /** Absent means "no harvest available" (tests, and a deliberate capture-only
   *  mint when M365_PERSONAL_MINT_STRATEGY=capture). */
  harvest?: (prev: PersonalRecord) => Promise<{ capture: PersonalCapture; rotated: boolean } | null>;
  now: () => number;
}

const DEFAULT_MINT_MIN_INTERVAL_MS = 600_000;
const DEFAULT_MINT_MAX_PER_HOUR = 3;
const HOUR = 3_600_000;

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
  // "No browser on :9222" and "the profile is dead" both used to advise `login`,
  // which is wrong for the first: login cannot start the operator's Chrome with
  // a debugging port, so the advice was unactionable (issue.md, agg. 1).
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

/** The CDP endpoint the mint path talks to: the active account's, so a mint
 *  can never open (or attach to) another account's Chrome (§7.2). */
export function probeCdpTarget(): string {
  return resolveAccount().cdpUrl;
}

/** The exact command that opens the active account's Chrome with the right
 *  debugging port. It used to hard-code 9222, which was the multi-account bug
 *  §7.2 names: a second account's Chrome could never be reached. */
export function chromeCommandLine(): string {
  const account = resolveAccount();
  return `chrome --remote-debugging-port=${new URL(account.cdpUrl).port || "9222"} --user-data-dir="${account.profileDir}" https://m365.cloud.microsoft/chat`;
}

/** A CDP endpoint that refused, or never answered, is not a dead session and
 *  not a bad capture: nothing was ever reached. */
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
  return new PersonalAuthError(
    "mint_rate_limited",
    `refusing to mint a fresh personal token: ${what} (the record is stale but not dead). Use M365_PERSONAL_MINT_MIN_INTERVAL_MS / M365_PERSONAL_MINT_MAX_PER_HOUR to tune, or ${personalLoginCommand()} to mint by hand.`,
  );
}

function defaultDeps(): PersonalAuthDeps {
  return {
    store: new PersonalStore(personalSessionFile()),
    capture: (headed) => capturePersonalSession({ headed, profileDir: personalProfileDir() }),
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

/** Which mechanism a re-mint uses. `harvest` (default) attaches to the
 *  operator's running Chrome and forces a socket; `capture` launches a browser
 *  like `login` does. Capture cannot pass Cloudflare itself (F101), so it is
 *  only a deliberate opt-out. */
function mintStrategy(): MintStrategy {
  return process.env.M365_PERSONAL_MINT_STRATEGY === "capture" ? "capture" : "harvest";
}

// The mint mutex (spec §10.2/§10.5): one browser per process. Reading the
// store is never held behind it.
let inflight: Promise<string> | null = null;

/** The token for personal mode: the stored one while fresh, a headless mint
 *  (then a headed one) when it is stale, and a rejected promise naming the
 *  login command when the profile is dead. */
export async function getPersonalToken(deps?: Partial<PersonalAuthDeps>): Promise<string> {
  const d: PersonalAuthDeps = { ...defaultDeps(), ...deps };
  const state = inspectQuietly(d.store);
  const record = state.ok ? state.record : null;
  if (record && d.store.isFresh(record, d.now())) return record.accessToken;

  // A record that exists but cannot be parsed is not "never logged in". The mint
  // still runs — it overwrites the bad file with a good one, which is the right
  // repair — but the operator should learn the file was broken, not lose that
  // fact in a silent overwrite. `error`, not `debug`: a corrupt record is a real
  // fault worth seeing without M365_DEBUG.
  if (!state.ok && state.problem !== "absent") {
    log.error(`the stored personal session is unusable (${state.problem}: ${state.detail}); re-minting over it`);
  }

  inflight ??= mint(d, record).finally(() => {
    inflight = null;
  });
  return inflight;
}

/** `inspect()` never throws by design; this also survives a stub store in tests
 *  that only implements `read()`. */
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
  // Another process may have minted while we waited on the mutex.
  const current = d.store.read();
  if (current && current !== prev && d.store.isFresh(current, d.now())) return current.accessToken;

  const now = d.now();
  const minInterval = Number(process.env.M365_PERSONAL_MINT_MIN_INTERVAL_MS ?? DEFAULT_MINT_MIN_INTERVAL_MS);
  const maxPerHour = Number(process.env.M365_PERSONAL_MINT_MAX_PER_HOUR ?? DEFAULT_MINT_MAX_PER_HOUR);
  if (prev) {
    if (now - prev.lastMintAt < minInterval) {
      log.debug(`mint refused: only ${Math.round((now - prev.lastMintAt) / 1000)}s since the last mint`);
      throw mintRateLimitedError("interval");
    }
    const mintsLastHour = now - prev.lastMintAt < HOUR ? prev.mintCount : 0;
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
    // A harvest spends a turn whether or not the token changed — a send usually
    // yields a new token (5/6 rounds, F111) but not always. Budgeting the send,
    // not the rotation, is what keeps the account from being hammered.
    mintCount: (prev ? (now - prev.lastMintAt < HOUR ? prev.mintCount : 0) : 0) + (spentTurn ? 1 : 0),
    lastMintAt: spentTurn ? mintedAt : (prev?.lastMintAt ?? mintedAt),
  };
  d.store.write(record);
  log.debug(`minted, expires in ${Math.round((record.expiresAt - mintedAt) / 60_000)} min`);
  return record.accessToken;
}

/** Harvest first, launch-capture second. The harvest is the path that works for
 *  a Chrome the operator left running; the capture is the old behaviour and is
 *  still the only thing that can mint when nothing is attached. */
async function obtainToken(
  d: PersonalAuthDeps,
  prev: PersonalRecord | null,
): Promise<{ capture: PersonalCapture; spentTurn: boolean }> {
  // No `prev` is not a reason to skip the harvest: the harvest needs a Chrome,
  // not a stored record. Requiring one made a corrupt record fatal — the harvest
  // was skipped, the capture path ran, Cloudflare refused it, and the operator
  // got `session_dead` for a file that only needed replacing.
  if (d.harvest) {
    try {
      const result = await d.harvest(prev ?? NO_PREVIOUS_RECORD);
      if (result) {
        if (!result.rotated) {
          // The page handed back the token we already had. It happened once
          // (two sends 3 min apart, F111); usually a send does produce a new
          // token. Keep the old one — it is not provably dead.
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

/** Stands in for the previous record when none could be read — a first mint, or
 *  a corrupt file. The harvest only reads `accessToken` from it (to tell whether
 *  the page rotated), so an empty token is enough: whatever the page hands back
 *  counts as new. */
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

/** The old mint: launch a browser, headless then headed. Kept intact for
 *  `login`-shaped mints and for a host with no attachable Chrome. */
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
      // A second launch cannot help when nothing is listening on the CDP port:
      // both attempts failed to connect, so report the port, not the profile.
      throw captureFailedError(isUnreachable(headReason) && isUnreachable(headedReason) ? headedReason : "");
    }
  }
}

/** Map a harvest failure onto the public error taxonomy. Returns null for an
 *  error the harvest layer did not originate (the caller falls back). */
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

/** The pre-flight `handleChatCompletion` awaits in personal mode: cached
 *  across turns so a healthy session costs nothing, and re-thrown (not
 *  cached) on failure so the 401 carries the real error each time. */
let cachedCredentials: Promise<void> | null = null;

export function ensurePersonalCredentials(): Promise<void> {
  cachedCredentials ??= getPersonalToken().then(
    () => undefined,
    (err) => {
      cachedCredentials = null;
      throw err;
    },
  );
  return cachedCredentials;
}
