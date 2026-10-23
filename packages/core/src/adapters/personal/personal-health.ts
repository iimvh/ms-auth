import { createLogger } from "../../log.js";
import { personalCdpUrl } from "./personal-harvest-cdp.js";
import type { PersonalRecord, PersonalRecordProblem, PersonalRecordState } from "./personal-session.js";
import { PersonalStore, personalSessionFile } from "./personal-session.js";

const log = createLogger("personal-health");

/** How a personal session looks from the outside, for `/healthz`. No turn is
 *  spent and no token is minted: this only reads the record and probes CDP. */
export interface PersonalHealth {
  /** `ok` — a fresh record and a reachable Chrome. `stale` — the record is past
   *  its (estimated) expiry but a mint can run. `at_risk` — one of the two is
   *  missing, so the next stale record becomes a hard failure. `corrupt` — a
   *  record file EXISTS but cannot be used, which is a different repair (delete
   *  it and log in again) from "never logged in". */
  status: "ok" | "stale" | "at_risk" | "corrupt";
  /** ms until the record's estimated expiry; negative once overdue. null when no
   *  usable record is stored. */
  expiresInMs: number | null;
  /** Whether the record passes the staleness policy right now. */
  fresh: boolean;
  /** Whether a Chrome with a debugging port answered. */
  cdpReachable: boolean;
  /** Present when status is not ok: what the operator has to do. */
  hint?: string;
  /** Present when the stored record could not be used: what is wrong with it. */
  problem?: PersonalRecordProblem;
}

export interface PersonalHealthDeps {
  store: Pick<PersonalStore, "inspect" | "isFresh">;
  probeCdp: (url: string) => Promise<boolean>;
  now: () => number;
}

/** Probe the CDP endpoint. A refused connection is a normal answer, not an
 *  error: it means "the operator's Chrome is not running". `url` names one
 *  account's Chrome explicitly; callers that have no account pass the active
 *  one. */
export async function probeCdpReachable(url: string, timeoutMs = 1000): Promise<boolean> {
  try {
    const res = await fetch(`${url.replace(/\/+$/, "")}/json/version`, {
      signal: AbortSignal.timeout(timeoutMs),
    });
    return res.ok;
  } catch (err) {
    log.debug(`CDP probe of ${url} failed:`, err);
    return false;
  }
}

function defaultDeps(): PersonalHealthDeps {
  return {
    store: new PersonalStore(personalSessionFile()),
    probeCdp: (url) => probeCdpReachable(url),
    now: () => Date.now(),
  };
}
/** Read the personal session's health. Never throws: a broken record is an
 *  unhealthy answer, not a 500. */
export async function personalHealth(deps?: Partial<PersonalHealthDeps>): Promise<PersonalHealth> {
  const d: PersonalHealthDeps = { ...defaultDeps(), ...deps };
  const now = d.now();
  let state: PersonalRecordState;
  try {
    state = d.store.inspect();
  } catch (err) {
    log.debug("inspecting the personal record failed:", err);
    state = {
      ok: false,
      problem: "unreadable",
      detail: err instanceof Error ? err.message : String(err),
    };
  }
  const cdpReachable = await d.probeCdp(personalCdpUrl());

  if (!state.ok) {
    // A record that is THERE but unusable is not "never logged in": saying so
    // would send the operator to `login` with no idea the old file is in the
    // way. `absent` keeps the plain login advice.
    if (state.problem === "absent") {
      return {
        status: "at_risk",
        expiresInMs: null,
        fresh: false,
        cdpReachable,
        problem: state.problem,
        hint: "no personal session stored — run: pnpm m365-personal login",
      };
    }
    return {
      status: "corrupt",
      expiresInMs: null,
      fresh: false,
      cdpReachable,
      problem: state.problem,
      hint: `the personal session record is unusable (${state.detail}) — delete it (or pnpm m365-personal logout) and run: pnpm m365-personal login`,
    };
  }
  const record = state.record;

  const expiresInMs = record.expiresAt - now;
  const fresh = d.store.isFresh(record, now);

  if (fresh) {
    return {
      status: "ok",
      expiresInMs,
      fresh,
      cdpReachable,
      // A fresh record with no Chrome is fine until the record goes stale, so
      // this is the one place we say so without calling it at_risk.
      ...(cdpReachable ? {} : { hint: "token is still fresh; start Chrome before it expires or the next mint cannot run" }),
    };
  }

  if (cdpReachable) {
    return {
      status: "stale",
      expiresInMs,
      fresh,
      cdpReachable,
      hint: "the record is past its estimated expiry — the next request will re-mint from the running Chrome",
    };
  }

  return {
    status: "at_risk",
    expiresInMs,
    fresh,
    cdpReachable,
    hint: `the record is past its estimated expiry and no Chrome is listening on ${personalCdpUrl()} — start it on the capture profile, or run: pnpm m365-personal login`,
  };
}
