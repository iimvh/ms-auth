import { createLogger } from "../../io/log.js";
import { PersonalCaptureError } from "./personal-capture.js";
import type { PersonalCapture } from "./personal-session.js";

const log = createLogger("personal-harvest");

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

const HARVEST_PROMPT = process.env.M365_PERSONAL_MINT_PROMPT ?? ".";

const DEFAULT_SOCKET_WAIT_MS = 20_000;

const CHATHUB_HOST = "substrate.office.com";
const CHATHUB_PATH_PREFIX = "/m365Copilot/Chathub/";

export interface HarvestedSocket {
  url: string;
  frames: string[];
}

export interface HarvestDeps {
  attach(): Promise<HarvestedPage>;
  now(): number;
}

export interface HarvestedPage {
  url(): string;
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

export function tokenFromSocketUrl(url: string): string | null {
  try {
    return new URL(url).searchParams.get("access_token") ?? null;
  } catch {
    return null;
  }
}

export interface HarvestResult {
  capture: PersonalCapture;
  rotated: boolean;
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

    const reused = deps.reuse?.(socket.url) ?? null;
    const capture: PersonalCapture = reused
      ? { ...reused, accessToken: token, capturedAt: deps.now() }
      : parseSocketCapture(socket.url, socket.frames, deps.now());
    return { capture, rotated, spentTurn: true };
  } finally {
    await page.close().catch((err: unknown) => log.debug("closing the harvest page handle failed:", err));
  }
}

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
    expiresAt: now + Number(process.env.M365_PERSONAL_TOKEN_TTL_MS ?? 3_600_000),
    capturedAt: now,
  };
}

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

export function isHarvestError(err: unknown): err is HarvestError {
  return err instanceof HarvestError;
}

export { PersonalCaptureError };
