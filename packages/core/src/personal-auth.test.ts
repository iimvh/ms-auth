import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { ensurePersonalCredentials, getPersonalToken, PersonalAuthError, resetPersonalAuthState } from "./personal-auth.js";
import { PersonalCaptureError } from "./personal-capture.js";
import { HarvestError } from "./personal-harvest.js";
import { PersonalStore, type PersonalRecord } from "./personal-session.js";

const HOUR = 3_600_000;
const fresh = (over: Partial<PersonalRecord> = {}): PersonalRecord => ({
  accessToken: "tok", chathubPath: "oid@tid", scenario: "S", variants: "v", licenseType: "Consumer",
  agentHost: "Bizchat.FullScreen", product: "Office", source: "officeweb", optionsSets: [],
  allowedMessageTypes: ["Chat"], expiresAt: 10 * HOUR, capturedAt: HOUR, emailHint: null, probe: null,
  mintCount: 1, lastMintAt: 0, ...over,
});
const setup = (record: PersonalRecord | null, over: Partial<{ harvest: unknown; capture: unknown }> = {}) => {
  const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
  if (record) store.write(record);
  const capture = vi.fn(async () => ({ ...fresh({ accessToken: "tok-new" }) }));
  // Harvesting is the default mint, so a test that does not want it must say so:
  // returning null is "no harvest available, capture instead".
  const harvest = vi.fn(async () => ({ capture: { ...fresh({ accessToken: "tok-harvested" }) }, rotated: true }));
  return {
    store,
    capture,
    harvest,
    deps: { store, capture, harvest, now: () => 2 * HOUR, ...over },
  };
};

/** A record whose harvest is deliberately unavailable. */
const captureOnly = (record: PersonalRecord | null) => setup(record, { harvest: vi.fn(async () => null) });

describe("PersonalAuthError", () => {
  it("is an Error whose code names the failure class", () => {
    const e = new PersonalAuthError("session_dead", "boom");
    expect(e).toBeInstanceOf(Error);
    expect(e.code).toBe("session_dead");
  });
});

describe("getPersonalToken", () => {
  it("returns a fresh token without minting", async () => {
    const { capture, harvest, deps } = setup(fresh());
    expect(await getPersonalToken(deps)).toBe("tok");
    expect(capture).not.toHaveBeenCalled();
    expect(harvest).not.toHaveBeenCalled();
  });

  it("harvests for a stale record and persists the result", async () => {
    const { store, capture, deps } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    expect(await getPersonalToken(deps)).toBe("tok-harvested");
    expect(capture).not.toHaveBeenCalled();
    expect(store.read()?.accessToken).toBe("tok-harvested");
  });

  it("two concurrent calls on a stale record mint once", async () => {
    const { harvest, deps } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    await Promise.all([getPersonalToken(deps), getPersonalToken(deps)]);
    expect(harvest).toHaveBeenCalledTimes(1);
  });

  it("a dead profile is session_dead and names the login command", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const deps = {
      store,
      now: () => 2 * HOUR,
      harvest: vi.fn(async () => null),
      capture: vi.fn(async () => {
        throw new PersonalCaptureError("not_chathub");
      }),
    };
    const err = await getPersonalToken(deps).catch((e) => e);
    expect(err).toBeInstanceOf(PersonalAuthError);
    expect(err.code).toBe("session_dead");
    expect(err.message).toContain("pnpm m365-personal login");
  });

  it("a corrupt record is minted over and replaced with a good one", async () => {
    // Minting over a broken file is the right repair — it leaves a usable record.
    // The point of `inspect()` is that the operator hears about the broken file
    // (personal-auth logs an error); it must not stop recovery.
    const { store, capture, deps } = captureOnly(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    writeFileSync(store.path, "{broken");
    expect(await getPersonalToken(deps)).toBe("tok-new");
    expect(capture).toHaveBeenCalledTimes(1);
    expect(store.read()?.accessToken).toBe("tok-new");
  });

  it("a corrupt record still harvests, instead of falling through to the dead capture path", async () => {
    // Regression: requiring a previous record to harvest made a corrupt file
    // fatal. The harvest is skipped, the capture runs, Cloudflare refuses it and
    // the operator sees `session_dead` for a file that only needed replacing.
    const { store, harvest } = setup(null);
    writeFileSync(store.path, "{broken");
    const capture = vi.fn(async () => {
      throw new PersonalCaptureError("capture_failed");
    });
    const deps = { store, capture, harvest, now: () => 2 * HOUR };
    expect(await getPersonalToken(deps)).toBe("tok-harvested");
    expect(harvest).toHaveBeenCalledTimes(1);
    expect(capture).not.toHaveBeenCalled();
  });

  it("retries a non-fatal capture failure once, headed", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const capture = vi.fn()
      .mockRejectedValueOnce(new PersonalCaptureError("capture_failed"))
      .mockResolvedValueOnce({ ...fresh({ accessToken: "tok-headless" }) });
    const token = await getPersonalToken({
      store,
      capture,
      harvest: vi.fn(async () => null),
      now: () => 2 * HOUR,
    });
    expect(token).toBe("tok-headless");
    expect(capture.mock.calls.map(([headed]) => headed)).toEqual([false, true]);
  });

  it("refuses a second mint inside the minimum interval", async () => {
    const { store, capture, harvest } = setup(
      fresh({ expiresAt: HOUR, capturedAt: 0, lastMintAt: 2 * HOUR - 1000, mintCount: 1 }),
    );
    const err = await getPersonalToken({ store, capture, harvest, now: () => 2 * HOUR }).catch((e) => e);
    expect(err.code).toBe("mint_rate_limited");
    expect(capture).not.toHaveBeenCalled();
    expect(harvest).not.toHaveBeenCalled();
  });
});

describe("a mint that cannot reach Chrome", () => {
  // Both attempts failing to CONNECT is not a dead profile and not a bad
  // capture: nothing was reached. The old message advised `login`, which cannot
  // start the operator's Chrome with a debugging port (issue.md aggravator 1).
  const noChrome = () =>
    vi.fn(async () => {
      throw Object.assign(new Error("connect ECONNREFUSED 127.0.0.1:9222"), { code: "ECONNREFUSED" });
    });

  it("is chrome_unreachable and names the port and the command", async () => {
    const { store } = captureOnly(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const err = await getPersonalToken({ store, capture: noChrome(), harvest: async () => null, now: () => 2 * HOUR }).catch((e) => e);
    expect(err).toBeInstanceOf(PersonalAuthError);
    expect(err.code).toBe("chrome_unreachable");
    expect(err.message).toContain("9222");
    expect(err.message).toContain("--remote-debugging-port");
    expect(err.message).not.toContain("pnpm m365-personal login");
  });

  it("still reports capture_failed when the browser was reached and failed later", async () => {
    const { store } = captureOnly(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const capture = vi.fn(async () => {
      throw new PersonalCaptureError("no_chat_frame");
    });
    const err = await getPersonalToken({ store, capture, harvest: async () => null, now: () => 2 * HOUR }).catch((e) => e);
    expect(err.code).toBe("capture_failed");
    expect(err.message).toContain("pnpm m365-personal login");
  });
});

describe("a harvest, which is the default mint", () => {
  // A send usually yields a new token (5/6 rounds, F111) but not always, so the
  // mint has to cope with both outcomes rather than assume either.
  it("keeps the token and does not charge a mint when the page did not rotate", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0, lastMintAt: HOUR, mintCount: 1 }));
    const harvest = vi.fn(async () => ({ capture: { ...fresh({ accessToken: "tok" }) }, rotated: false }));
    const token = await getPersonalToken({ store, capture: vi.fn(), harvest, now: () => 2 * HOUR });
    expect(token).toBe("tok");
  });

  it("falls back to a capture when no harvest is available", async () => {
    const { store, capture } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const token = await getPersonalToken({ store, capture, harvest: async () => null, now: () => 2 * HOUR });
    expect(token).toBe("tok-new");
    expect(capture).toHaveBeenCalledTimes(1);
  });

  it("an unreachable Chrome does NOT silently fall back to a launch attempt", async () => {
    const { store, capture } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const harvest = vi.fn(async () => {
      throw new HarvestError("chrome_unreachable", "connect ECONNREFUSED");
    });
    const err = await getPersonalToken({ store, capture, harvest, now: () => 2 * HOUR }).catch((e) => e);
    expect(err.code).toBe("chrome_unreachable");
    expect(capture).not.toHaveBeenCalled();
  });

  it("no chat tab maps to no_chat_tab", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const harvest = vi.fn(async () => {
      throw new HarvestError("no_chat_tab", "no m365 page");
    });
    const err = await getPersonalToken({ store, harvest, now: () => 2 * HOUR }).catch((e) => e);
    expect(err.code).toBe("no_chat_tab");
    expect(err.message).toContain("m365.cloud.microsoft/chat");
  });
});

describe("per-account mint state", () => {
  it("two accounts mint in parallel, one mutex each", async () => {
    const a = setup(fresh({ accessToken: "a-old", expiresAt: HOUR, capturedAt: 0 }));
    const b = setup(fresh({ accessToken: "b-old", expiresAt: HOUR, capturedAt: 0 }));
    b.harvest.mockImplementation(async () => ({ capture: { ...fresh({ accessToken: "b-new" }) }, rotated: true }));
    const [ta, tb] = await Promise.all([
      getPersonalToken({ ...a.deps, accountId: "a" }),
      getPersonalToken({ ...b.deps, accountId: "b" }),
    ]);
    expect(ta).toBe("tok-harvested");
    expect(tb).toBe("b-new");
    expect(a.harvest).toHaveBeenCalledTimes(1);
    expect(b.harvest).toHaveBeenCalledTimes(1);
  });

  it("the pre-flight cache does not serve one account's success to another", async () => {
    resetPersonalAuthState();
    const ok = vi.fn(async () => "tok");
    await ensurePersonalCredentials("a", { getToken: ok });
    const boom = vi.fn(async () => { throw new Error("account b is dead"); });
    await expect(ensurePersonalCredentials("b", { getToken: boom })).rejects.toThrow(/account b is dead/);
    expect(boom).toHaveBeenCalledTimes(1);
    await ensurePersonalCredentials("a", { getToken: ok });
    expect(ok).toHaveBeenCalledTimes(1); // a's success is still cached
  });

  it("a failure is not cached, so the next call reports it again", async () => {
    resetPersonalAuthState();
    const boom = vi.fn(async () => { throw new Error("nope"); });
    await expect(ensurePersonalCredentials("c", { getToken: boom })).rejects.toThrow();
    await expect(ensurePersonalCredentials("c", { getToken: boom })).rejects.toThrow();
    expect(boom).toHaveBeenCalledTimes(2);
  });
});