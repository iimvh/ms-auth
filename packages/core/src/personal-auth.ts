import { capturePersonalSession, PersonalCaptureError } from "./personal-capture.js";
import type { PersonalCapture, PersonalRecord } from "./personal-session.js";
import { PersonalStore, personalProfileDir, personalSessionFile } from "./personal-session.js";
import { createLogger } from "./log.js";

const log = createLogger("personal-auth");

export type PersonalAuthErrorCode = "session_dead" | "mint_rate_limited" | "capture_failed";

export class PersonalAuthError extends Error {
  constructor(
    readonly code: PersonalAuthErrorCode,
    message: string,
  ) {
    super(message);
    this.name = "PersonalAuthError";
  }
}

export interface PersonalAuthDeps {
  store: PersonalStore;
  capture: (headed: boolean) => Promise<PersonalCapture>;
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

function captureFailedError(): PersonalAuthError {
  return new PersonalAuthError(
    "capture_failed",
    `could not re-mint the M365 personal session — run: ${personalLoginCommand()}`,
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
    now: () => Date.now(),
  };
}

// The mint mutex (spec §10.2/§10.5): one browser per process. Reading the
// store is never held behind it.
let inflight: Promise<string> | null = null;

/** The token for personal mode: the stored one while fresh, a headless mint
 *  (then a headed one) when it is stale, and a rejected promise naming the
 *  login command when the profile is dead. */
export async function getPersonalToken(deps?: Partial<PersonalAuthDeps>): Promise<string> {
  const d: PersonalAuthDeps = { ...defaultDeps(), ...deps };
  const record = d.store.read();
  if (record && d.store.isFresh(record, d.now())) return record.accessToken;

  inflight ??= mint(d, record).finally(() => {
    inflight = null;
  });
  return inflight;
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

  log.debug("mint attempt (headless)");
  let capture: PersonalCapture;
  try {
    capture = await d.capture(false);
  } catch (err) {
    if (err instanceof PersonalCaptureError && err.reason === "not_chathub") {
      log.debug(`mint failed: the profile is dead (${err.message})`);
      throw sessionDeadError();
    }
    log.debug(`mint failed headlessly: ${describeCaptureError(err)}; retrying headed`);
    try {
      capture = await d.capture(true);
    } catch (retryErr) {
      log.debug(`mint failed headed too: ${describeCaptureError(retryErr)}`);
      throw captureFailedError();
    }
  }

  const mintedAt = d.now();
  const record: PersonalRecord = {
    ...capture,
    emailHint: prev?.emailHint ?? null,
    probe: prev?.probe ?? null,
    mintCount: (prev ? (now - prev.lastMintAt < HOUR ? prev.mintCount : 0) : 0) + 1,
    lastMintAt: mintedAt,
  };
  d.store.write(record);
  log.debug(`minted, expires in ${Math.round((record.expiresAt - mintedAt) / 60_000)} min`);
  return record.accessToken;
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
