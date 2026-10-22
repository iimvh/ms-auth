import { chmodSync, mkdirSync, readFileSync, renameSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { createLogger } from "./log.js";

const log = createLogger("personal-session");

const CONFIG_DIR = join(homedir(), ".config", "opencode-m365");

function resolveFile(envVar: string, defaultName: string): string {
  if (process.env[envVar]) return process.env[envVar]!;
  mkdirSync(CONFIG_DIR, { recursive: true });
  return join(CONFIG_DIR, defaultName);
}

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

/** The `personal-session.json` cache. Re-reads on every use — the cache key
 *  is the raw file bytes (mtime+size alone hides sub-tick changes on
 *  Windows), so a `login` in another process takes effect without a restart
 *  (spec §10.2). Writes are temp-then-rename; the record carries a bearer
 *  token, so POSIX files get 0o600 (Windows ACL hardening is §9). */
export class PersonalStore {
  readonly path: string;
  private cache: PersonalRecord | null = null;
  private cacheKey: string | null = null;

  constructor(file: string) {
    this.path = file;
  }

  read(): PersonalRecord | null {
    let raw: string;
    let key: string;
    try {
      raw = readFileSync(this.path, "utf8");
      const s = statSync(this.path);
      const st = { mtimeMs: s.mtimeMs, size: s.size };
      // Windows papers over sub-tick writes (two renames inside one timestamp
      // tick), so mtime+size alone can hide a change; the raw bytes are the
      // authoritative key. The stat pair is only a debug-friendly hint.
      key = raw;
      void st;
    } catch {
      this.cache = null;
      this.cacheKey = null;
      return null;
    }
    if (this.cache && this.cacheKey === key) return this.cache;
    this.cacheKey = key;
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch (err) {
      log.debug("personal record is not valid JSON, treating as absent:", err);
      this.cache = null;
      return null;
    }
    if (!isPersonalRecord(parsed)) {
      log.debug("personal record failed the shape check, treating as absent");
      this.cache = null;
      return null;
    }
    this.cache = parsed;
    return this.cache;
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
  }

  delete(): void {
    try {
      unlinkSync(this.path);
    } catch (err) {
      log.debug("delete of the personal record failed (ok if absent):", err);
    }
    this.cache = null;
    this.cacheKey = null;
  }

  isFresh(record: PersonalRecord, now: number = Date.now()): boolean {
    const lifetime = record.expiresAt - record.capturedAt;
    const skew = Math.min(Math.max(300_000, 0.1 * lifetime), 0.5 * lifetime);
    return now < record.expiresAt - skew;
  }
}

export function personalSessionFile(): string {
  return resolveFile("M365_PERSONAL_SESSION_FILE", "personal-session.json");
}

export function personalProfileDir(): string {
  return resolveFile("M365_PERSONAL_PROFILE", "personal-profile");
}
