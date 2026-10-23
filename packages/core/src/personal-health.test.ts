import { describe, expect, it, vi } from "vitest";
import { personalHealth } from "./personal-health.js";
import { PersonalStore, type PersonalRecord } from "./personal-session.js";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const HOUR = 3_600_000;

const record = (over: Partial<PersonalRecord> = {}): PersonalRecord => ({
  accessToken: "tok",
  chathubPath: "oid@tid",
  scenario: "S",
  variants: "v",
  licenseType: "Free",
  agentHost: "Bizchat.FullScreen",
  product: null,
  source: "officeweb",
  optionsSets: [],
  allowedMessageTypes: ["Chat"],
  expiresAt: 4 * HOUR,
  capturedAt: 2 * HOUR,
  emailHint: null,
  probe: null,
  mintCount: 1,
  lastMintAt: 2 * HOUR,
  ...over,
});

/** A store backed by a real file (the health path must not care how it reads). */
function storeWith(rec: PersonalRecord | null): PersonalStore {
  const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ph-")), "s.json"));
  if (rec) store.write(rec);
  return store;
}

const deps = (rec: PersonalRecord | null, cdp: boolean) => ({
  store: storeWith(rec),
  probeCdp: vi.fn(async () => cdp),
  now: () => 3 * HOUR,
});

describe("personalHealth", () => {
  it("ok when the record is fresh and Chrome is reachable", async () => {
    const h = await personalHealth(deps(record(), true));
    expect(h.status).toBe("ok");
    expect(h.fresh).toBe(true);
    expect(h.cdpReachable).toBe(true);
    expect(h.hint).toBeUndefined();
  });

  it("ok but hints when the record is fresh and no Chrome is running", async () => {
    // Not at_risk: a fresh token serves turns without a browser. The risk is
    // what happens when it goes stale, so the hint says that.
    const h = await personalHealth(deps(record(), false));
    expect(h.status).toBe("ok");
    expect(h.cdpReachable).toBe(false);
    expect(h.hint).toContain("start Chrome");
  });

  it("stale when the record is expired and Chrome can re-mint it", async () => {
    const h = await personalHealth(deps(record({ expiresAt: 2 * HOUR + 60_000 }), true));
    expect(h.status).toBe("stale");
    expect(h.fresh).toBe(false);
    expect(h.expiresInMs).toBeLessThan(0);
    expect(h.hint).toContain("re-mint");
  });

  it("at_risk when the record is expired and no Chrome is listening", async () => {
    const h = await personalHealth(deps(record({ expiresAt: 2 * HOUR + 60_000 }), false));
    expect(h.status).toBe("at_risk");
    expect(h.hint).toContain("no Chrome is listening");
    expect(h.hint).toContain("m365-personal login");
  });

  it("at_risk with no record at all", async () => {
    const h = await personalHealth(deps(null, true));
    expect(h.status).toBe("at_risk");
    expect(h.expiresInMs).toBeNull();
    expect(h.hint).toContain("pnpm m365-personal login");
  });

  it("never throws, even when the store blows up", async () => {
    const store = {
      inspect: () => {
        throw new Error("disk on fire");
      },
      isFresh: () => false,
    } as unknown as PersonalStore;
    const h = await personalHealth({ store, probeCdp: async () => false, now: () => 0 });
    // Present-but-unreadable is `corrupt`, not `at_risk`: the repair differs.
    expect(h.status).toBe("corrupt");
    expect(h.problem).toBe("unreadable");
  });

  it("a record file that is present but broken is corrupt, not 'not logged in'", async () => {
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ph-")), "s.json"));
    writeFileSync(store.path, '{"accessToken": "tok"');
    const h = await personalHealth({ store, probeCdp: async () => true, now: () => 0 });
    expect(h.status).toBe("corrupt");
    expect(h.problem).toBe("invalid_json");
    expect(h.hint).toContain("delete it");
    expect(h.hint).not.toBe("no personal session stored — run: pnpm m365-personal login");
  });

  it("a well-formed file that is not a record says which fields are missing", async () => {
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ph-")), "s.json"));
    writeFileSync(store.path, JSON.stringify({ accessToken: "tok" }));
    const h = await personalHealth({ store, probeCdp: async () => true, now: () => 0 });
    expect(h.status).toBe("corrupt");
    expect(h.problem).toBe("wrong_shape");
    expect(h.hint).toContain("chathubPath");
  });

  it("a record saved with a UTF-8 BOM is still readable (not 'never logged in')", async () => {
    // PowerShell's `Set-Content -Encoding UTF8` and many Windows editors do this;
    // JSON.parse rejects the BOM, which used to make a good record look absent.
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ph-")), "s.json"));
    const good = record();
    writeFileSync(store.path, "\uFEFF" + JSON.stringify(good));
    const h = await personalHealth({ store, probeCdp: async () => true, now: () => 3 * HOUR });
    expect(h.status).toBe("ok");
    expect(store.read()?.accessToken).toBe("tok");
  });

  it("spends no turn: it never touches the mint path", async () => {
    // Structural guarantee: the only deps are read/isFresh and a CDP probe.
    const d = deps(record({ expiresAt: 0 }), false);
    await personalHealth(d);
    expect(d.probeCdp).toHaveBeenCalledTimes(1);
  });
});
