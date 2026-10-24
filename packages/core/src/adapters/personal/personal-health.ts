import { createLogger } from "../../io/log.js";
import { personalCdpUrl } from "./personal-harvest-cdp.js";
import type { PersonalRecord, PersonalRecordProblem, PersonalRecordState } from "./personal-session.js";
import { PersonalStore, personalSessionFile } from "./personal-session.js";

const log = createLogger("personal-health");

export interface PersonalHealth {
  status: "ok" | "stale" | "at_risk" | "corrupt";
  expiresInMs: number | null;
  fresh: boolean;
  cdpReachable: boolean;
  hint?: string;
  problem?: PersonalRecordProblem;
}

export interface PersonalHealthDeps {
  store: Pick<PersonalStore, "inspect" | "isFresh">;
  probeCdp: (url: string) => Promise<boolean>;
  now: () => number;
}

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
