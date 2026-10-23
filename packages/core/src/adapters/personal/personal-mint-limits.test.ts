import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import {
  PERSONAL_AUTH_ERROR_CODES,
  PersonalAuthError,
  getPersonalToken,
  type PersonalAuthErrorCode,
} from "./personal-auth.js";
import {
  MINT_HOUR_MS,
  PERSONAL_MINT_RATE_LIMIT_ENV,
  PERSONAL_MINT_RATE_LIMITS,
  PersonalMintRateLimitOverrideSchema,
  personalMintRateLimits,
} from "./personal-mint-limits.js";
import { PersonalStore, type PersonalRecord } from "./personal-session.js";

const fresh = (over: Partial<PersonalRecord> = {}): PersonalRecord => ({
  accessToken: "tok", chathubPath: "oid@tid", scenario: "S", variants: "v", licenseType: "Consumer",
  agentHost: "Bizchat.FullScreen", product: "Office", source: "officeweb", optionsSets: [],
  allowedMessageTypes: ["Chat"], expiresAt: 10 * MINT_HOUR_MS, capturedAt: MINT_HOUR_MS, emailHint: null, probe: null,
  mintCount: 1, lastMintAt: 0, ...over,
});

const staleStore = (over: Partial<PersonalRecord> = {}): PersonalStore => {
  const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
  store.write(fresh({ expiresAt: MINT_HOUR_MS, capturedAt: 0, ...over }));
  return store;
};

const mintAt = (store: PersonalStore, now: number, rateLimits = {}) => {
  const harvest = vi.fn(async () => ({ capture: { ...fresh({ accessToken: "tok-new" }) }, rotated: true }));
  return {
    store,
    harvest,
    token: getPersonalToken({ store, harvest, capture: vi.fn(), now: () => now, rateLimits }),
  };
};

describe("mint rate limits", () => {
  it("default to a ten-minute minimum interval and three mints an hour", () => {
    expect(personalMintRateLimits({})).toEqual({ minIntervalMs: 600_000, maxPerHour: 3 });
    expect(PERSONAL_MINT_RATE_LIMITS).toEqual({ minIntervalMs: 600_000, maxPerHour: 3 });
    expect(PERSONAL_MINT_RATE_LIMIT_ENV).toEqual({
      minIntervalMs: "M365_PERSONAL_MINT_MIN_INTERVAL_MS",
      maxPerHour: "M365_PERSONAL_MINT_MAX_PER_HOUR",
    });
    expect(MINT_HOUR_MS).toBe(3_600_000);
  });

  it("take the value each named env variable carries", () => {
    expect(
      personalMintRateLimits({
        [PERSONAL_MINT_RATE_LIMIT_ENV.minIntervalMs]: "60000",
        [PERSONAL_MINT_RATE_LIMIT_ENV.maxPerHour]: "9",
      }),
    ).toEqual({ minIntervalMs: 60_000, maxPerHour: 9 });
  });

  it("fall back to the default for a value that is not a number", () => {
    expect(PERSONAL_MINT_RATE_LIMITS).toEqual({ minIntervalMs: 600_000, maxPerHour: 3 });
    expect(personalMintRateLimits({ [PERSONAL_MINT_RATE_LIMIT_ENV.minIntervalMs]: "soon" })).toEqual({
      minIntervalMs: 600_000,
      maxPerHour: 3,
    });
    expect(personalMintRateLimits({ [PERSONAL_MINT_RATE_LIMIT_ENV.maxPerHour]: "many" })).toEqual({
      minIntervalMs: 600_000,
      maxPerHour: 3,
    });
  });

  it("accept an explicit override ahead of the environment, and reject a negative one", () => {
    const env = { [PERSONAL_MINT_RATE_LIMIT_ENV.minIntervalMs]: "1", [PERSONAL_MINT_RATE_LIMIT_ENV.maxPerHour]: "1" };
    expect(personalMintRateLimits(env, { minIntervalMs: 5, maxPerHour: 6 })).toEqual({ minIntervalMs: 5, maxPerHour: 6 });
    expect(PersonalMintRateLimitOverrideSchema.safeParse({ maxPerHour: -1 }).success).toBe(false);
    expect(() => personalMintRateLimits({}, { minIntervalMs: -1 })).toThrow();
  });

  it("refuse a mint inside the minimum interval, and allow one exactly at it", async () => {
    const refused = mintAt(staleStore({ lastMintAt: 2 * MINT_HOUR_MS - 1 }), 2 * MINT_HOUR_MS);
    await expect(refused.token).rejects.toMatchObject({ code: "mint_rate_limited" });
    expect(refused.harvest).not.toHaveBeenCalled();

    const boundary = mintAt(staleStore({ lastMintAt: 2 * MINT_HOUR_MS - PERSONAL_MINT_RATE_LIMITS.minIntervalMs }), 2 * MINT_HOUR_MS);
    await expect(boundary.token).resolves.toBe("tok-new");
    expect(boundary.harvest).toHaveBeenCalledTimes(1);
  });

  it("refuse the mint after the third in an hour, but only while that mint is inside the hour", async () => {
    const insideTheHour = mintAt(staleStore({ lastMintAt: 2 * MINT_HOUR_MS - 1, mintCount: 3 }), 2 * MINT_HOUR_MS);
    await expect(insideTheHour.token).rejects.toMatchObject({ code: "mint_rate_limited" });
    expect(insideTheHour.harvest).not.toHaveBeenCalled();

    const atTheHour = mintAt(staleStore({ lastMintAt: MINT_HOUR_MS, mintCount: 3 }), 2 * MINT_HOUR_MS);
    await expect(atTheHour.token).resolves.toBe("tok-new");
    expect(atTheHour.store.read()).toMatchObject({ mintCount: 1, lastMintAt: 2 * MINT_HOUR_MS });
  });

  it("count the mints inside the hour only, and charge one per harvest", async () => {
    const inside = staleStore({ lastMintAt: 2 * MINT_HOUR_MS - PERSONAL_MINT_RATE_LIMITS.minIntervalMs, mintCount: 2 });
    await mintAt(inside, 2 * MINT_HOUR_MS).token;
    expect(inside.read()).toMatchObject({ mintCount: 3, lastMintAt: 2 * MINT_HOUR_MS });

    const carriedOver = staleStore({ lastMintAt: MINT_HOUR_MS - 1, mintCount: 5 });
    await mintAt(carriedOver, 2 * MINT_HOUR_MS).token;
    expect(carriedOver.read()?.mintCount).toBe(1);
  });

  it("charge nothing to a mint that falls back to a capture, which spends no turn", async () => {
    const lastMintAt = 2 * MINT_HOUR_MS - PERSONAL_MINT_RATE_LIMITS.minIntervalMs;
    const store = staleStore({ lastMintAt, mintCount: 2 });
    const capture = vi.fn(async () => ({ ...fresh({ accessToken: "tok-captured" }) }));
    const token = await getPersonalToken({ store, capture, harvest: async () => null, now: () => 2 * MINT_HOUR_MS });
    expect(token).toBe("tok-captured");
    expect(store.read()).toMatchObject({ mintCount: 2, lastMintAt });
  });
});

describe("the personal auth error codes", () => {
  const ALL: PersonalAuthErrorCode[] = [
    "session_dead",
    "mint_rate_limited",
    "capture_failed",
    "chrome_unreachable",
    "no_chat_tab",
    "mint_not_rotated",
  ];

  it("are exactly the six a PersonalAuthError may carry", () => {
    expect([...PERSONAL_AUTH_ERROR_CODES].sort()).toEqual([...ALL].sort());
    for (const code of ALL) expect(new PersonalAuthError(code, "boom").code).toBe(code);
  });
});
