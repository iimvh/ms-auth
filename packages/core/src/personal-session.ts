import { chmodSync, mkdirSync, readFileSync, renameSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { resolveAccount } from "./account-registry.js";
import { createLogger } from "./log.js";

const log = createLogger("personal-session");

/** One turn's worth of captured connection parameters, exactly as the real
 *  m365.cloud.microsoft/chat client sent them (§6.1 of the personal-auth spec). */
export interface PersonalCapture {
  accessToken: string;
  chathubPath: string;
  scenario: string;
  variants: string;
  licenseType: string;
  agentHost: string;
  /** The paid work surface carries `product=Office`; the free consumer socket
   *  sends no `product` at all (F97) — replay must omit it then, so no fake
   *  value is invented here. */
  product: string | null;
  source: string;
  optionsSets: string[];
  allowedMessageTypes: string[];
  expiresAt: number;
  capturedAt: number;
}

/** What the record file holds: the capture plus bookkeeping for mint limits,
 *  the login-time model probe, and the account hint. */
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

/** The fields `isPersonalRecord` insists on, named for the operator's benefit:
 *  "missing expiresAt" is actionable, "wrong shape" is not. */
function missingRecordFields(value: unknown): string[] {
  if (typeof value !== "object" || value === null) return ["the whole record (not an object)"];
  const r = value as Record<string, unknown>;
  const problems: string[] = [];
  if (typeof r.accessToken !== "string") problems.push("accessToken");
  if (typeof r.chathubPath !== "string") problems.push("chathubPath");
  if (typeof r.expiresAt !== "number") problems.push("expiresAt");
  return problems;
}

/** Why a stored record could not be used. `absent` is the ordinary "nothing has
 *  been captured yet"; the rest mean a file IS there but cannot be trusted, and
 *  telling them apart matters: `absent` is fixed by `login`, a corrupt file is a
 *  bug or a bad hand-edit and `login` overwrites evidence of it. */
export type PersonalRecordProblem =
  | "absent"
  | "unreadable"
  | "invalid_json"
  | "wrong_shape";

/** What `PersonalStore.inspect()` found: the record when usable, otherwise the
 *  reason and enough detail to say something useful. */
export type PersonalRecordState =
  | { ok: true; record: PersonalRecord }
  | { ok: false; problem: PersonalRecordProblem; detail: string };

/** Strip a UTF-8 byte-order mark. Windows editors (and PowerShell's
 *  `Set-Content -Encoding UTF8`) write one, and `JSON.parse` rejects it — a
 *  perfectly good record would otherwise read as "no session stored". */
export function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** The `personal-session.json` cache. Re-reads on every use — the cache key
 *  is the raw file bytes (mtime+size alone hides sub-tick changes on
 *  Windows), so a `login` in another process takes effect without a restart
 *  (spec §10.2). Writes are temp-then-rename; the record carries a bearer
 *  token, so POSIX files get 0o600 (Windows ACL hardening is §9). */
export class PersonalStore {
  readonly path: string;
  private cache: PersonalRecord | null = null;
  private cacheKey: string | null = null;
  private state: PersonalRecordState | null = null;

  constructor(file: string) {
    this.path = file;
  }

  /** The record, or null when there is nothing usable. Callers that need to
   *  distinguish "never captured" from "captured but broken" use `inspect()`. */
  read(): PersonalRecord | null {
    const state = this.inspect();
    return state.ok ? state.record : null;
  }

  /** The record plus, when there is no usable record, why. Never throws. */
  inspect(): PersonalRecordState {
    let raw: string;
    try {
      raw = readFileSync(this.path, "utf8");
      // Windows papers over sub-tick writes (two renames inside one timestamp
      // tick), so mtime+size alone can hide a change; the raw bytes are the
      // authoritative key.
    } catch (err) {
      const missing = (err as NodeJS.ErrnoException)?.code === "ENOENT";
      if (missing) {
        this.clearCache();
        return this.setState({ ok: false, problem: "absent", detail: "no personal session file" });
      }
      // Present but unreadable (permissions, a directory, a locked file).
      this.clearCache();
      return this.setState({
        ok: false,
        problem: "unreadable",
        detail: `${this.path}: ${err instanceof Error ? err.message : String(err)}`,
      });
    }

    // The cache key includes the BOM state, so a file rewritten with one is a
    // different entry even when the JSON text is otherwise identical.
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
    // The state cache is the one `inspect()` serves; leaving it stale here would
    // hand the caller the record we just replaced.
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
