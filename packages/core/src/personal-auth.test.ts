import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { getPersonalToken, PersonalAuthError } from "./personal-auth.js";
import { PersonalCaptureError } from "./personal-capture.js";
import { PersonalStore, type PersonalRecord } from "./personal-session.js";

const HOUR = 3_600_000;
const fresh = (over: Partial<PersonalRecord> = {}): PersonalRecord => ({
  accessToken: "tok", chathubPath: "oid@tid", scenario: "S", variants: "v", licenseType: "Consumer",
  agentHost: "Bizchat.FullScreen", product: "Office", source: "officeweb", optionsSets: [],
  allowedMessageTypes: ["Chat"], expiresAt: 10 * HOUR, capturedAt: HOUR, emailHint: null, probe: null,
  mintCount: 1, lastMintAt: 0, ...over,
});
const setup = (record: PersonalRecord | null) => {
  const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
  if (record) store.write(record);
  const capture = vi.fn(async () => ({ ...fresh({ accessToken: "tok-new" }) }));
  return { store, capture, deps: { store, capture, now: () => 2 * HOUR } };
};

describe("PersonalAuthError", () => {
  it("is an Error whose code names the failure class", () => {
    const e = new PersonalAuthError("session_dead", "boom");
    expect(e).toBeInstanceOf(Error);
    expect(e.code).toBe("session_dead");
  });
});

describe("getPersonalToken", () => {
  it("returns a fresh token without capturing", async () => {
    const { store, capture, deps } = setup(fresh());
    expect(await getPersonalToken(deps)).toBe("tok");
    expect(capture).not.toHaveBeenCalled();
  });

  it("mints once for an expired record and persists the result", async () => {
    const { store, capture, deps } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    expect(await getPersonalToken(deps)).toBe("tok-new");
    expect(capture).toHaveBeenCalledTimes(1);
    expect(store.read()?.accessToken).toBe("tok-new");
  });

  it("two concurrent calls on a stale record mint once", async () => {
    const { capture, deps } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    await Promise.all([getPersonalToken(deps), getPersonalToken(deps)]);
    expect(capture).toHaveBeenCalledTimes(1);
  });

  it("a dead profile is session_dead and names the login command", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const deps = { store, now: () => 2 * HOUR,
      capture: vi.fn(async () => { throw new PersonalCaptureError("not_chathub"); }) };
    const err = await getPersonalToken(deps).catch((e) => e);
    expect(err).toBeInstanceOf(PersonalAuthError);
    expect(err.code).toBe("session_dead");
    expect(err.message).toContain("pnpm m365-personal login");
  });

  it("a corrupt record is treated as absent and minted over", async () => {
    const { store, capture, deps } = setup(null);
    writeFileSync(store.path, "{broken");
    expect(await getPersonalToken(deps)).toBe("tok-new");
    expect(capture).toHaveBeenCalledTimes(1);
  });

  it("retries a non-fatal capture failure once, headed", async () => {
    const { store } = setup(fresh({ expiresAt: HOUR, capturedAt: 0 }));
    const capture = vi.fn()
      .mockRejectedValueOnce(new PersonalCaptureError("capture_failed"))
      .mockResolvedValueOnce({ ...fresh({ accessToken: "tok-headless" }) });
    const token = await getPersonalToken({ store, capture, now: () => 2 * HOUR });
    expect(token).toBe("tok-headless");
    expect(capture.mock.calls.map(([headed]) => headed)).toEqual([false, true]);
  });

  it("refuses a second mint inside the minimum interval", async () => {
    const { store, capture } = setup(fresh({ expiresAt: HOUR, capturedAt: 0, lastMintAt: 2 * HOUR - 1000, mintCount: 1 }));
    const err = await getPersonalToken({ store, capture, now: () => 2 * HOUR }).catch((e) => e);
    expect(err.code).toBe("mint_rate_limited");
    expect(capture).not.toHaveBeenCalled();
  });
});
