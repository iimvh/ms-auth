import { existsSync, mkdtempSync, readdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PersonalStore, type PersonalRecord } from "./personal-session.js";

const record = (over: Partial<PersonalRecord> = {}): PersonalRecord => ({
  accessToken: "tok-1", chathubPath: "oid@tid", scenario: "OfficeWebPaidConsumerCopilot",
  variants: "a,b", licenseType: "Consumer", agentHost: "Bizchat.FullScreen",
  product: "Office", source: "officeweb", optionsSets: ["enable_msa_user"],
  allowedMessageTypes: ["Chat"], expiresAt: 1_000_000, capturedAt: 0,
  emailHint: null, probe: null, mintCount: 1, lastMintAt: 0, ...over,
});

describe("PersonalStore", () => {
  it("round-trips a record and leaves no temp file behind", () => {
    const dir = mkdtempSync(join(tmpdir(), "ps-"));
    const store = new PersonalStore(join(dir, "s.json"));
    store.write(record());
    expect(store.read()?.accessToken).toBe("tok-1");
    expect(readdirSync(dir)).toEqual(["s.json"]);
  });

  it("returns null for a missing file", () => {
    expect(new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json")).read()).toBeNull();
  });

  it("treats truncated JSON or a missing field as absent instead of throwing", () => {
    const dir = mkdtempSync(join(tmpdir(), "ps-"));
    const file = join(dir, "s.json");
    const store = new PersonalStore(file);
    writeFileSync(file, '{"accessToken":"tok"');
    expect(store.read()).toBeNull();
    writeFileSync(file, JSON.stringify({ accessToken: "tok" }));
    expect(store.read()).toBeNull();
  });

  it("sees a write made by another process on the next read", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ps-")), "s.json");
    const reader = new PersonalStore(file);
    reader.write(record());
    expect(reader.read()?.accessToken).toBe("tok-1");
    new PersonalStore(file).write(record({ accessToken: "tok-2", capturedAt: 5 }));
    expect(reader.read()?.accessToken).toBe("tok-2");
  });

  it("is fresh only outside a skew of max(5 min, 10% of lifetime)", () => {
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
    const hour = 3_600_000;
    expect(store.isFresh(record({ capturedAt: 0, expiresAt: hour }), 500_000)).toBe(true);
    expect(store.isFresh(record({ capturedAt: 0, expiresAt: hour }), hour - 350_000)).toBe(false);
    // a 60 s token gets a 30 s skew, not a 5 min one that would always be stale
    expect(store.isFresh(record({ capturedAt: 0, expiresAt: 60_000 }), 20_000)).toBe(true);
  });

  it("deletes the file", () => {
    const file = join(mkdtempSync(join(tmpdir(), "ps-")), "s.json");
    const store = new PersonalStore(file);
    store.write(record());
    store.delete();
    expect(existsSync(file)).toBe(false);
  });
});
