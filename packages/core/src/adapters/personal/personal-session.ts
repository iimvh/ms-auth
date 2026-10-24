import { chmodSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { z } from "zod/v4";
import { resolveAccount } from "../../account-registry.js";
import { createLogger } from "../../io/log.js";

const log = createLogger("personal-session");

export interface PersonalCapture {
  accessToken: string;
  chathubPath: string;
  scenario: string;
  variants: string;
  licenseType: string;
  agentHost: string;
  product: string | null;
  source: string;
  optionsSets: string[];
  allowedMessageTypes: string[];
  expiresAt: number;
  capturedAt: number;
}

export interface PersonalRecord extends PersonalCapture {
  emailHint: string | null;
  probe: { model: string; ok: boolean; at: number } | null;
  mintCount: number;
  lastMintAt: number;
}

function isPersonalRecord(value: unknown): value is PersonalRecord {
  if (typeof value !== "object" || value === null) return false;
  const r = value as Record<string, unknown>;
  return (
    typeof r.accessToken === "string" &&
    typeof r.chathubPath === "string" &&
    typeof r.expiresAt === "number"
  );
}

function missingRecordFields(value: unknown): string[] {
  if (typeof value !== "object" || value === null) return ["the whole record (not an object)"];
  const r = value as Record<string, unknown>;
  const problems: string[] = [];
  if (typeof r.accessToken !== "string") problems.push("accessToken");
  if (typeof r.chathubPath !== "string") problems.push("chathubPath");
  if (typeof r.expiresAt !== "number") problems.push("expiresAt");
  return problems;
}

export const PersonalCaptureSchema = z.object({
  accessToken: z.string(),
  chathubPath: z.string(),
  scenario: z.string(),
  variants: z.string(),
  licenseType: z.string(),
  agentHost: z.string(),
  product: z.string().nullable(),
  source: z.string(),
  optionsSets: z.array(z.string()),
  allowedMessageTypes: z.array(z.string()),
  expiresAt: z.number(),
  capturedAt: z.number(),
});

export const PersonalRecordSchema = PersonalCaptureSchema.extend({
  emailHint: z.string().nullable(),
  probe: z.object({ model: z.string(), ok: z.boolean(), at: z.number() }).nullable(),
  mintCount: z.number(),
  lastMintAt: z.number(),
});

export const PersonalRecordProblemSchema = z.enum(["absent", "unreadable", "invalid_json", "wrong_shape"]);

export type PersonalRecordProblem =
  | "absent"
  | "unreadable"
  | "invalid_json"
  | "wrong_shape";

export type PersonalRecordState =
  | { ok: true; record: PersonalRecord }
  | { ok: false; problem: PersonalRecordProblem; detail: string };

export function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

export class PersonalStore {
  readonly path: string;
  private cache: PersonalRecord | null = null;
  private cacheKey: string | null = null;
  private state: PersonalRecordState | null = null;

  constructor(file: string) {
    this.path = file;
  }

  read(): PersonalRecord | null {
    const state = this.inspect();
    return state.ok ? state.record : null;
  }

  inspect(): PersonalRecordState {
    let raw: string;
    try {
      raw = readFileSync(this.path, "utf8");
    } catch (err) {
      const missing = (err as NodeJS.ErrnoException)?.code === "ENOENT";
      if (missing) {
        this.clearCache();
        return this.setState({ ok: false, problem: "absent", detail: "no personal session file" });
      }
      this.clearCache();
      return this.setState({
        ok: false,
        problem: "unreadable",
        detail: `${this.path}: ${err instanceof Error ? err.message : String(err)}`,
      });
    }

    if (this.state && this.cacheKey === raw) return this.state;

    const text = stripBom(raw);
    let parsed: unknown;
    try {
      parsed = JSON.parse(text);
    } catch (err) {
      log.debug("personal record is not valid JSON:", err);
      this.cache = null;
      this.cacheKey = raw;
      const hint = raw.charCodeAt(0) === 0xfeff ? " (a UTF-8 BOM is handled; the rest is not JSON)" : "";
      return this.setState({
        ok: false,
        problem: "invalid_json",
        detail: `${this.path} is not valid JSON${hint}: ${err instanceof Error ? err.message : String(err)}`,
      });
    }
    if (!isPersonalRecord(parsed)) {
      log.debug("personal record failed the shape check");
      this.cache = null;
      this.cacheKey = raw;
      const missing = missingRecordFields(parsed);
      return this.setState({
        ok: false,
        problem: "wrong_shape",
        detail: `${this.path} is not a personal session record (missing or invalid: ${missing.join(", ") || "unknown"})`,
      });
    }
    this.cache = parsed;
    this.cacheKey = raw;
    return this.setState({ ok: true, record: parsed });
  }

  private clearCache(): void {
    this.cache = null;
    this.cacheKey = null;
  }

  private setState(state: PersonalRecordState): PersonalRecordState {
    this.state = state;
    return state;
  }

  write(record: PersonalRecord): void {
    const dir = dirname(this.path);
    mkdirSync(dir, { recursive: true });
    const tmp = `${this.path}.tmp-${process.pid}`;
    const raw = JSON.stringify(record, null, 2) + "\n";
    writeFileSync(tmp, raw, { mode: 0o600 });
    renameSync(tmp, this.path);
    if (process.platform !== "win32") {
      try {
        chmodSync(this.path, 0o600);
      } catch (err) {
        log.debug("chmod 0o600 on the personal record failed:", err);
      }
    }
    this.cache = record;
    this.cacheKey = raw;
    this.state = { ok: true, record };
  }

  delete(): void {
    try {
      unlinkSync(this.path);
    } catch (err) {
      log.debug("delete of the personal record failed (ok if absent):", err);
    }
    this.clearCache();
    this.state = { ok: false, problem: "absent", detail: "no personal session file" };
  }

  isFresh(record: PersonalRecord, now: number = Date.now()): boolean {
    const lifetime = record.expiresAt - record.capturedAt;
    const skew = Math.min(Math.max(300_000, 0.1 * lifetime), 0.5 * lifetime);
    return now < record.expiresAt - skew;
  }
}

export function personalSessionFile(): string {
  return resolveAccount().sessionFile;
}

export function personalProfileDir(): string {
  return resolveAccount().profileDir;
}
