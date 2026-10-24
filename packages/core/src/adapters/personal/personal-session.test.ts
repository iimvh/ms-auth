import { existsSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import {
  PersonalCaptureSchema,
  PersonalRecordProblemSchema,
  PersonalRecordSchema,
  PersonalStore,
  type PersonalRecord,
  type PersonalRecordState,
} from "./personal-session.js";

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

describe("a stored record's state", () => {
  it("has exactly the four problems, and each one is what inspect() reports for its file", () => {
    const dir = mkdtempSync(join(tmpdir(), "ps-"));
    const stateOf = (content: string | null, path = join(dir, "s.json")): PersonalRecordState => {
      if (content !== null) writeFileSync(path, content);
      return new PersonalStore(path).inspect();
    };
    const states: Record<string, PersonalRecordState> = {
      absent: stateOf(null, join(dir, "missing.json")),
      unreadable: stateOf(null, dir),
      invalid_json: stateOf('{"accessToken":"tok"'),
      wrong_shape: stateOf(JSON.stringify({ accessToken: "tok" })),
    };
    expect(Object.keys(states).sort()).toEqual([...PersonalRecordProblemSchema.options].sort());
    for (const [problem, state] of Object.entries(states)) {
      expect(state.ok).toBe(false);
      expect(state).toMatchObject({ problem });
    }
    const usable = stateOf(JSON.stringify(record()));
    expect(usable.ok).toBe(true);
    expect(usable).toMatchObject({ record: { accessToken: "tok-1" } });
  });
});

describe("PersonalRecordSchema", () => {
  it("accepts a written record and rejects each field whose type changes", () => {
    const dir = mkdtempSync(join(tmpdir(), "ps-"));
    const store = new PersonalStore(join(dir, "s.json"));
    store.write(record());
    const written = JSON.parse(readFileSync(store.path, "utf8")) as unknown;
    expect(PersonalRecordSchema.safeParse(written).success).toBe(true);
    expect(PersonalCaptureSchema.safeParse(record()).success).toBe(true);

    const broken: Record<keyof PersonalRecord, unknown> = {
      accessToken: 1, chathubPath: null, scenario: false, variants: [], licenseType: 0, agentHost: {},
      product: undefined, source: 1, optionsSets: "s", allowedMessageTypes: {}, expiresAt: "1e3",
      capturedAt: null, emailHint: 7, probe: { model: "m", ok: "yes", at: 0 }, mintCount: "1", lastMintAt: null,
    };
    for (const [field, value] of Object.entries(broken)) {
      expect({ field, accepted: PersonalRecordSchema.safeParse({ ...record(), [field]: value }).success }).toEqual({ field, accepted: false });
    }
  });
});
