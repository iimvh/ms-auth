import { createLogger } from "./log.js";
import { PersonalCaptureError } from "./personal-capture.js";
import type { PersonalCapture } from "./personal-session.js";

const log = createLogger("personal-harvest");

/** Why a harvest failed to produce a token. Same enum as the capture reasons
 *  plus the two states only an attached browser can be in. */
export type HarvestErrorReason =
  | "chrome_unreachable"
  | "no_chat_tab"
  | "session_dead"
  | "no_chat_frame"
  | "capture_failed";

export class HarvestError extends Error {
  constructor(
    readonly reason: HarvestErrorReason,
    detail?: string,
  ) {
    super(detail);
    this.name = "HarvestError";
  }
}

/** The message a harvest sends to force the Chathub socket open. It costs a turn
 *  (the free consumer account is capped at 80/day), so it is deliberately the
 *  shortest thing that still submits. */
const HARVEST_PROMPT = process.env.M365_PERSONAL_MINT_PROMPT ?? ".";

/** How long to wait for the socket after clicking Send. A send opens it ~1 s
 *  later (measured, F111); the slack covers a cold composer and a slow network. */
const DEFAULT_SOCKET_WAIT_MS = 20_000;

const CHATHUB_HOST = "substrate.office.com";
const CHATHUB_PATH_PREFIX = "/m365Copilot/Chathub/";

/** A websocket the page opened, reduced to what a harvest needs. */
export interface HarvestedSocket {
  url: string;
  frames: string[];
}

export interface HarvestDeps {
  /** Attach to an already-running Chrome and hand back its chat page. Throws
   *  `chrome_unreachable`/`no_chat_tab` when that is impossible. */
  attach(): Promise<HarvestedPage>;
  now(): number;
}

/** The slice of a Playwright page a harvest drives. Kept structural so the
 *  harvest is unit-testable without a browser. */
export interface HarvestedPage {
  url(): string;
  /** Open the Chathub socket by sending one message; resolves once a socket with
   *  a chat invocation frame is seen, or with `null` if none appeared. */
  sendAndWait(waitMs: number): Promise<HarvestedSocket | null>;
  close(): Promise<void>;
}

function isChathubUrl(url: string): boolean {
  try {
    const p = new URL(url);
    return p.host === CHATHUB_HOST && p.pathname.startsWith(CHATHUB_PATH_PREFIX);
  } catch {
    return false;
  }
}

/** The token out of a Chathub socket URL, or null. A harvest reads the token
 *  from the URL only — measured to exist nowhere in page storage (F111). */
export function tokenFromSocketUrl(url: string): string | null {
  try {
    return new URL(url).searchParams.get("access_token") ?? null;
  } catch {
    return null;
  }
}

export interface HarvestResult {
  capture: PersonalCapture;
  /** Whether the token differs from the one the caller already had. A harvest
   *  usually gets a new one (5 of 6 rounds a minute apart did), but it is not
   *  guaranteed: two sends 3 minutes apart once returned the same token. The
   *  honesty matters more than the expectation — see F111. */
  rotated: boolean;
  /** The turn spent on the send. Counted even when the token did not rotate. */
  spentTurn: boolean;
}

/** Force the page to open its Chathub socket, read the token off it, and decide
 *  whether that token is actually new.
 *
 *  Why sending is the mechanism: the page opens no socket on load and none on
 *  focus — only a send opens one, and it closes ~6 s later (F111). There is no
 *  long-lived connection to harvest passively and the token is in no storage. */
export async function harvestPersonalToken(deps: {
  attach: () => Promise<HarvestedPage>;
  reuse?: (socketUrl: string) => PersonalCapture | null;
  previousToken: string | null;
  waitMs?: number;
  now: () => number;
}): Promise<HarvestResult> {
  const waitMs = deps.waitMs ?? Number(process.env.M365_PERSONAL_MINT_SOCKET_WAIT_MS ?? DEFAULT_SOCKET_WAIT_MS);
  const page = await deps.attach();
  try {
    log.debug(`harvest: sending ${JSON.stringify(HARVEST_PROMPT)} to force the socket, waitMs=${waitMs}`);
    const socket = await page.sendAndWait(waitMs);
    if (!socket) {
      throw new HarvestError("no_chat_frame", "no Chathub socket opened after sending a message");
    }
    if (!isChathubUrl(socket.url)) {
      throw new HarvestError("capture_failed", `the socket opened was not a Chathub socket: ${socket.url.slice(0, 120)}`);
    }
    const token = tokenFromSocketUrl(socket.url);
    if (!token) {
      throw new HarvestError("capture_failed", "the Chathub socket URL carried no access_token");
    }
    const rotated = deps.previousToken === null || token !== deps.previousToken;
    log.debug(`harvest: token ${rotated ? "rotated" : "unchanged"} (len ${token.length})`);

    // When the token did not rotate the socket still carries the exact
    // connection params, so reuse them; only the token/lifetime is renewed.
    const reused = deps.reuse?.(socket.url) ?? null;
    const capture: PersonalCapture = reused
      ? { ...reused, accessToken: token, capturedAt: deps.now() }
      : parseSocketCapture(socket.url, socket.frames, deps.now());
    return { capture, rotated, spentTurn: true };
  } finally {
    await page.close().catch((err: unknown) => log.debug("closing the harvest page handle failed:", err));
  }
}

/** Rebuild a capture from a socket URL on the harvest path. The full parse lives
 *  in `personal-capture.ts`; this is only used when the caller could not supply
 *  the previous params. */
function parseSocketCapture(url: string, frames: string[], now: number): PersonalCapture {
  const parsed = new URL(url);
  const get = (name: string): string => {
    const v = parsed.searchParams.get(name);
    if (v === null || v === "") throw new HarvestError("capture_failed", `the socket URL is missing ${name}`);
    return v;
  };
  let source = parsed.searchParams.get("source") ?? "";
  try {
    source = JSON.parse(source) as string;
  } catch {
    // not JSON-quoted; use verbatim
  }
  return {
    accessToken: get("access_token"),
    chathubPath: parsed.pathname.slice(CHATHUB_PATH_PREFIX.length).replace(/^\/+|\/+$/g, ""),
    scenario: get("scenario"),
    variants: get("variants"),
    licenseType: get("licenseType"),
    agentHost: get("agentHost"),
    product: parsed.searchParams.get("product") ?? null,
    source,
    optionsSets: frameStringArray(frames, "optionsSets"),
    allowedMessageTypes: frameStringArray(frames, "allowedMessageTypes"),
    // The socket token is an opaque JWE (F106): keep the existing estimate.
    expiresAt: now + Number(process.env.M365_PERSONAL_TOKEN_TTL_MS ?? 3_600_000),
    capturedAt: now,
  };
}

/** Best-effort pull of a string[] field out of the chat invocation frame. */
function frameStringArray(frames: string[], field: string): string[] {
  for (const raw of frames) {
    const m = new RegExp(`"${field}"\\s*:\\s*\\[([^\\]]*)\\]`).exec(raw);
    if (!m) continue;
    const out: string[] = [];
    for (const piece of m[1]!.split(",")) {
      const v = piece.trim().replace(/^"|"$/g, "");
      if (v) out.push(v);
    }
    return out;
  }
  return [];
}

/** Re-export so callers can distinguish a harvest failure from a capture one. */
export function isHarvestError(err: unknown): err is HarvestError {
  return err instanceof HarvestError;
}

export { PersonalCaptureError };
