import { execSync } from "node:child_process";
import { resolveAccount } from "./account-registry.js";
import { createLogger } from "./log.js";
import type { PersonalCapture } from "./personal-session.js";

const log = createLogger("personal-capture");

export type PersonalCaptureErrorReason =
  | "not_chathub"
  | "no_chat_frame"
  | "not_a_jwt"
  | "missing_param"
  | "capture_failed";

export class PersonalCaptureError extends Error {
  constructor(
    readonly reason: PersonalCaptureErrorReason,
    detail?: string,
  ) {
    const what: Record<PersonalCaptureErrorReason, string> = {
      not_chathub: "the captured websocket is not a Chathub socket",
      no_chat_frame: "the client sent no chat invocation frame to capture",
      not_a_jwt: "the captured access_token is not a JWT (no exp claim readable)",
      missing_param: detail ?? "the captured websocket URL is missing a required parameter",
      capture_failed: detail ?? "the browser capture failed before a usable session was read",
    };
    // The detail is the diagnosis (e.g. "the browser landed on a sign-in page");
    // swallowing it leaves the operator with "not a Chathub socket" for a dead
    // profile. missing_param's detail IS its message.
    super(reason === "missing_param" ? what.missing_param : detail ? `${what[reason]}: ${detail}` : what[reason]);
    this.name = "PersonalCaptureError";
  }
}

const CHATHUB_HOST = "substrate.office.com";
const CHATHUB_PATH_PREFIX = "/m365Copilot/Chathub/";

/** Expiry of a captured access token, from its `exp` claim, in ms. */
export function decodeJwtExpiresAt(token: string): number {
  const parts = token.split(".");
  if (parts.length !== 3) {
    throw new PersonalCaptureError("not_a_jwt");
  }
  try {
    const payload = JSON.parse(Buffer.from(parts[1]!, "base64url").toString("utf8"));
    const exp = (payload as Record<string, unknown>).exp;
    if (typeof exp !== "number" || !Number.isFinite(exp)) {
      throw new PersonalCaptureError("not_a_jwt");
    }
    return exp * 1000;
  } catch (err) {
    if (err instanceof PersonalCaptureError) throw err;
    throw new PersonalCaptureError("not_a_jwt");
  }
}

/** Which URL params to read and whether the value is JSON-quoted by the
 *  client (`%22officeweb%22` → `officeweb`). `product` is optional: the paid
 *  work socket sends it, the free consumer socket does not (F97). */
const STRING_PARAMS: readonly string[] = [
  "scenario",
  "variants",
  "licenseType",
  "agentHost",
];
const OPTIONAL_STRING_PARAMS: readonly string[] = ["product"];
const JSON_PARAMS: readonly string[] = ["source"];

/** Rebuild a `PersonalCapture` from the real client's Chathub socket: the URL
 *  carries the connection params, the first sent `type:4, target:"chat"`
 *  invocation carries `optionsSets`/`allowedMessageTypes`, the JWT carries
 *  `exp` (spec §6.1). Invoked by the browser driver in Task 3. */
export function parsePersonalCapture(wsUrl: string, sentFrames: string[]): PersonalCapture {
  const url = new URL(wsUrl);
  if (url.host !== CHATHUB_HOST || !url.pathname.startsWith(CHATHUB_PATH_PREFIX)) {
    throw new PersonalCaptureError("not_chathub");
  }
  const chathubPath = url.pathname.slice(CHATHUB_PATH_PREFIX.length).replace(/^\/+|\/+$/g, "");

  const accessToken = url.searchParams.get("access_token");
  if (!accessToken) {
    throw new PersonalCaptureError("missing_param", "access_token");
  }
  let expiresAt: number;
  try {
    expiresAt = decodeJwtExpiresAt(accessToken);
  } catch (err) {
    // The consumer (free-tier) socket's token is an encrypted JWE (alg=dir,
    // A256CBC-HS512, zip=DEF) — its claims are not readable without Microsoft's
    // key (F98). Treat a non-JWT as opaque and estimate a TTL; the staleness
    // policy then mints early, and a wrong estimate only costs one mint.
    //
    // The default 1 h is NOT a measurement: a captured token kept serving turns
    // at least 77 minutes after capture (measured 2026-10-23), and the page's
    // own MSAL credential for the same account lives 8 h. The socket token's
    // real lifetime is unknown — see docs/personal-mint-design.md.
    if (!(err instanceof PersonalCaptureError) || err.reason !== "not_a_jwt") throw err;
    const ttl = Number(process.env.M365_PERSONAL_TOKEN_TTL_MS ?? 3_600_000);
    if (!Number.isFinite(ttl) || ttl <= 0) throw err;
    expiresAt = Date.now() + ttl;
    log.debug(`access_token is opaque (not a readable JWT) — expiry estimated from a ${ttl} ms TTL`);
  }

  const required = new Map<string, string>();
  for (const name of STRING_PARAMS) {
    const v = url.searchParams.get(name);
    if (v === null || v === "") {
      throw new PersonalCaptureError("missing_param", name);
    }
    required.set(name, v);
  }
  for (const name of JSON_PARAMS) {
    const raw = url.searchParams.get(name);
    if (raw === null || raw === "") {
      throw new PersonalCaptureError("missing_param", name);
    }
    let value = raw;
    try {
      value = JSON.parse(raw) as string;
    } catch {
      log.debug(`captured ${name}=${raw} was not JSON-quoted; using it verbatim`);
    }
    required.set(name, value);
  }

  const chatArg = firstChatArgument(sentFrames);
  const capture: PersonalCapture = {
    accessToken,
    chathubPath,
    scenario: required.get("scenario")!,
    variants: required.get("variants")!,
    licenseType: required.get("licenseType")!,
    agentHost: required.get("agentHost")!,
    product: url.searchParams.get("product") ?? null,
    source: required.get("source")!,
    optionsSets: stringArray(chatArg?.optionsSets),
    allowedMessageTypes: stringArray(chatArg?.allowedMessageTypes),
    expiresAt,
    capturedAt: Date.now(),
  };
  log.debug(
    `captured chathubPath=${chathubPath} scenario=${capture.scenario} licenseType=${capture.licenseType}`,
  );
  return capture;
}

const RS = "\x1e";

/** All visible text across the page and its frames (Cloudflare's challenge lives
 *  in an out-of-process iframe whose text the top document never shows). */
async function pageText(page: {
  locator(sel: string): { innerText(): Promise<string> };
  frames(): Array<{ locator(sel: string): { innerText(): Promise<string> } }>;
}): Promise<string> {
  const parts: string[] = [];
  for (const frame of page.frames()) {
    try {
      parts.push(await frame.locator("body").innerText());
    } catch {
      // detached/frame navigated away
    }
  }
  return parts.join(String.fromCharCode(10));
}

/** Whether the page is asking for Microsoft's human verification. Reading the
 *  sentence is all this does — nothing tries to solve or bypass the challenge. */
export function bodyNeedsVerification(text: string): boolean {
  // F100/F101: the gate surfaced as "Verification required to continue." on one
  // run and as a "Security check required" dialog followed by "Your request
  // couldn't be completed. Refresh to try again." on another — all three mean a
  // human has to act in the window.
  return /verification required to continue|verify you are human|security check required|couldn.t be completed\. refresh to try again/i.test(text);
}

/** Whether the sent frames include the chat invocation (type:4, target:"chat").
 *  A socket that has only opened and exchanged the SignalR handshake is NOT
 *  enough: the app opens the Chathub socket on page load and only sends the
 *  invocation when a message is actually sent. */
export function hasChatFrame(sentFrames: string[]): boolean {
  try {
    firstChatArgument(sentFrames);
    return true;
  } catch {
    return false;
  }
}

function firstChatArgument(sentFrames: string[]): Record<string, unknown> | null {
  for (const rawFrame of sentFrames) {
    for (const message of splitMessages(rawFrame)) {
      const f = message as Record<string, unknown>;
      if (f.type !== 4 || f.target !== "chat") continue;
      const args = f.arguments;
      if (Array.isArray(args) && args.length > 0 && typeof args[0] === "object" && args[0] !== null) {
        return args[0] as Record<string, unknown>;
      }
    }
  }
  throw new PersonalCaptureError("no_chat_frame");
}

/** One WS frame can carry several SignalR messages: normally separated by the
 *  record separator (), but the consumer client also concatenates them
 *  with no separator at all (F102 — a `type:4` chat invocation followed by a
 *  `Metrics` message, so a whole-frame JSON.parse fails halfway). Split on the
 *  separator, then on JSON object boundaries, and keep what parses. */
export function splitMessages(rawFrame: string): unknown[] {
  const messages: unknown[] = [];
  for (const chunk of rawFrame.split(RS)) {
    for (const piece of splitConcatenatedJson(chunk)) {
      try {
        messages.push(JSON.parse(piece));
      } catch {
        // not JSON (e.g. the Socket.IO-style "5:::" frames) — not ours
      }
    }
  }
  return messages;
}

/** Text that is one or more concatenated JSON objects → each object's source
 *  text. A string-aware brace counter: `}` inside a JSON string never closes. */
function splitConcatenatedJson(text: string): string[] {
  const pieces: string[] = [];
  let depth = 0;
  let start = -1;
  let inString = false;
  let escaped = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]!;
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') inString = true;
    else if (ch === "{") {
      if (depth === 0) start = i;
      depth++;
    } else if (ch === "}") {
      depth--;
      if (depth === 0 && start >= 0) {
        pieces.push(text.slice(start, i + 1));
        start = -1;
      }
    }
  }
  return pieces;
}

function stringArray(value: unknown): string[] {
  if (value === undefined || value === null) return [];
  if (Array.isArray(value)) return value.filter((v): v is string => typeof v === "string");
  return [];
}

/** The message used to force a Chathub socket open when the app opens none on
 *  page load. Whatever the model replies, it is not recorded — only the socket
 *  the reply travelled on is. */
const PROMPT_TO_TRIGGER_SOCKET = "hi";

function findChromium(): string | undefined {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;
  for (const bin of ["chromium", "chromium-browser", "google-chrome", "chrome"]) {
    try {
      const found = execSync(`command -v ${bin}`, { stdio: ["ignore", "pipe", "ignore"] })
        .toString()
        .trim();
      if (found) return found;
    } catch {
      // not on PATH, try next
    }
  }
  return undefined;
}

const LOGIN_USER_AGENT =
  process.env.M365_LOGIN_UA ??
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36";

const CHAT_HOME = "https://m365.cloud.microsoft/chat";

/** A URL that is (part of) the AAD sign-in flow rather than the chat app. */
export function isLoginUrl(url: string): boolean {
  return url.includes("login.microsoftonline.com") || url.includes("/login");
}

/** Stored email/password/TOTP when the operator has them — the same
 *  `secrets.json` the enterprise automated login uses. Unset/absent/short is
 *  simply "no automation, wait for a human". */
async function storedSecrets(): Promise<{ email: string; password: string; mfaSecret: string } | null> {
  if (process.env.M365_PERSONAL_AUTOMATED_LOGIN === "0") return null;
  try {
    const { loadSecrets } = await import("./auth.js");
    return loadSecrets();
  } catch {
    return null;
  }
}

/** A WS payload as text: text frames arrive as strings; some clients send the
 *  same JSON as binary frames — decode UTF-8 and keep it if it reads like a
 *  JSON frame. Playwright's payload type has grown over versions (string,
 *  Buffer, Uint8Array, plain object); none of this is worth a crash: the
 *  handler catches everything and ignores what it cannot read. */
export function frameText(data: unknown): string | null {
  try {
    // Playwright >=1.5x hands `framesent` an object `{ payload }`; older code
    // paths hand the payload itself. Accept both.
    const inner =
      data && typeof data === "object" && "payload" in (data as Record<string, unknown>)
        ? (data as { payload: unknown }).payload
        : data;
    if (typeof inner === "string") return inner;
    if (inner instanceof ArrayBuffer) return maybeJson(new Uint8Array(inner));
    if (inner instanceof Uint8Array) return maybeJson(inner);
    return null;
  } catch {
    return null;
  }
}

function maybeJson(bytes: Uint8Array): string | null {
  const text = Buffer.from(bytes).toString("utf8");
  return text.startsWith("{") ? text : null;
}

/** Drive the real m365.cloud.microsoft/chat in a persistent profile and
 *  capture the exact Chathub connection the app opens (spec §6.1). Headless
 *  for a mint, headed for the once-a-day `login`. This is operator-facing
 *  code — no unit test; the parse it feeds is tested in Task 2. */
export async function capturePersonalSession(opts: {
  headed: boolean;
  profileDir: string;
  waitMs?: number;
  /** Operator-facing progress lines (e.g. "sign in in the opened window").
   *  The CLI wires it to stdout; a mint leaves it unset (debug file only). */
  onEvent?: (line: string) => void;
}): Promise<PersonalCapture> {  const waitMs = opts.waitMs ?? Number(process.env.M365_PERSONAL_MINT_WAIT_MS ?? 10_000);
  log.debug(`capturing personal session (${opts.headed ? "headed" : "headless"}), waitMs=${waitMs}`);
  const { chromium } = await import("playwright");
  // Cloudflare's "verify you are human" gate refuses a Playwright-launched
  // browser no matter how the window is dressed: the launch itself is the
  // fingerprint. With a registry (an active account) or M365_PERSONAL_CDP_URL
  // set, the operator starts a REAL Chrome (`--remote-debugging-port=<port>
  // --user-data-dir=<profile dir>`) and we attach over CDP instead — nothing to
  // launch, no automation flag (F101). The account's own CDP target is what
  // keeps a capture from ever landing in another account's profile.
  const resolved = resolveAccount();
  const cdpUrl = resolved.id !== null || process.env.M365_PERSONAL_CDP_URL ? resolved.cdpUrl : null;
  const attached = cdpUrl ? await chromium.connectOverCDP(cdpUrl) : null;
  const context = attached
    ? attached.contexts()[0] ?? (await attached.newContext())
    : await chromium.launchPersistentContext(opts.profileDir, {
        headless: !opts.headed,
        executablePath: findChromium(),
        args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
        userAgent: LOGIN_USER_AGENT,
        locale: process.env.M365_LOGIN_LOCALE ?? "en-GB",
        timezoneId: process.env.M365_LOGIN_TIMEZONE ?? "Europe/Copenhagen",
        viewport: { width: 1280, height: 800 },
      });
  if (attached) log.debug(`attached over CDP to ${cdpUrl} — the operator's own Chrome`);
  try {
    const page =
      context.pages().find((p) => p.url().includes("m365.cloud.microsoft")) ??
      context.pages()[0] ??
      (await context.newPage());
    await page.addInitScript(() => {
      Object.defineProperty(navigator, "webdriver", { get: () => undefined });
    });

    const sockets: { url: string; frames: string[] }[] = [];
    page.on("websocket", (ws) => {
      log.debug(`ws opened: ${ws.url()}`);
      const rec = { url: ws.url(), frames: [] as string[] };
      sockets.push(rec);
      ws.on("framesent", (payload) => {
        const text = frameText(payload);
        if (text) {
          log.debug(`frame sent (${text.length} chars): ${text.length > 300 ? text.slice(0, 300) : text}`);
          rec.frames.push(text);
        } else {
          log.debug(`non-JSON or unreadable frame ignored`);
        }
      });
    });

    await page.goto(CHAT_HOME, { waitUntil: "domcontentloaded" });
    log.debug(`landing page: ${page.url()}`);
    // A consumer sign-in always transits an AAD authorize URL (the /common/
    // oauth2 endpoint) on the way to the chat app — even with a valid saved
    // session, it auto-continues there in a couple of seconds. Treat
    // "landing on a login URL" as a state to out-wait, not a fact: a signed-in
    // profile leaves it quickly, an unsigned one stays. Headed enrollment then
    // keeps waiting for the operator (up to loginWaitMs); a headless mint
    // gives the silent SSO only a bounded transit window (transitMs) before
    // declaring the profile dead (spec §6.3).
    if (isLoginUrl(page.url()) && opts.headed) {
      const loginWaitMs = Number(process.env.M365_PERSONAL_LOGIN_WAIT_MS ?? 300_000);
      opts.onEvent?.(
        `the browser opened on the Microsoft sign-in page — complete the sign-in there; waiting up to ${Math.round(loginWaitMs / 1000)} s`,
      );
      log.debug(`headed login waits for interactive sign-in, loginWaitMs=${loginWaitMs}`);
      // AGENTS auth order: silent → automated (secrets.json) → ask a human. The
      // visible window stays open, so both paths race and whichever finishes first
      // wins; automation needs M365_PERSONAL_AUTOMATED_LOGIN=0 to opt out.
      const creds = await storedSecrets();
      if (creds) {
        opts.onEvent?.("trying the stored credentials' automated sign-in");
        const { driveSignInForm } = await import("./auth.js");
        void driveSignInForm(page, creds)
          .then(() => log.debug("automated sign-in finished"))
          .catch((err) => log.debug("automated sign-in ended:", err?.message ?? err));
      }
      const deadline = Date.now() + loginWaitMs;
      while (Date.now() < deadline && isLoginUrl(page.url())) {
        await page.waitForTimeout(500);
      }
    } else if (isLoginUrl(page.url())) {
      const transitMs = Number(process.env.M365_PERSONAL_TRANSIT_WAIT_MS ?? 30_000);
      log.debug(`headless: giving the silent SSO ${transitMs} ms to leave the authorize URL`);
      const deadline = Date.now() + transitMs;
      while (Date.now() < deadline && isLoginUrl(page.url())) {
        await page.waitForTimeout(500);
      }
    }
    if (isLoginUrl(page.url())) {
      log.debug(`still on a login page after the wait: ${page.url()}`);
      throw new PersonalCaptureError(
        "not_chathub",
        opts.headed
          ? "the browser stayed on a sign-in page, so the profile's session is dead"
          : "the browser landed on a sign-in page, so the profile's session is dead",
      );
    }

    const isChathub = (u: string) => {
      try {
        const p = new URL(u);
        return p.host === CHATHUB_HOST && p.pathname.startsWith(CHATHUB_PATH_PREFIX);
      } catch {
        return false;
      }
    };

    const waitForSocket = async (ms: number) => {
      const deadline = Date.now() + ms;
      while (Date.now() < deadline) {
        // The socket alone is not the prize: the capture needs the chat
        // invocation frame (optionsSets/allowedMessageTypes). A page-load
        // Chathub socket that only exchanges the SignalR handshake leaves
        // firstChatArgument with nothing (F99), so typed-prompt forcing must
        // still happen in that case.
        const socket = sockets.find((s) => isChathub(s.url) && hasChatFrame(s.frames));
        if (socket) return socket;
        await page.waitForTimeout(250);
      }
      return null;
    };

    // Page-load capture first (spec §10.3): typing anything costs a thread.
    let socket = await waitForSocket(waitMs);
    if (!socket) {
      log.debug(`no chat invocation frame after ${waitMs} ms; typing a short message to force one`);
      // The consumer page's editor is a contenteditable span; typing through the
      // keyboard (not fill) is what lets its editor state and the Send button
      // react (F100). The Send button is preferred over Enter for the same reason.
      const composer = page
        .locator('#m365-chat-editor-target-element, textarea, [role="textbox"], [contenteditable="true"]')
        .first();
      await composer.click({ timeout: 30_000 });
      // Clear any saved draft first: the page restores unsent text across runs,
      // and typing on top of it produced "kkkhi" instead of "hi" (F101).
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(PROMPT_TO_TRIGGER_SOCKET, { delay: 40 });
      const send = page.locator('button[aria-label*="Send" i], button[title*="Send" i]').first();
      if (await send.count()) await send.click({ timeout: 10_000 });
      else await page.keyboard.press("Enter");
      socket = await waitForSocket(waitMs);
      if (!socket) {
        // Microsoft gates a free consumer conversation behind a one-time human
        // verification ("Verification required to continue.") — the message is
        // never sent, so no invocation frame can appear (F100). Wait for the
        // operator to clear it in this window, then send again. Nothing here
        // solves or bypasses the challenge.
        // The gate dialog can lag the send by a few seconds; give it a beat.
        // Cloudflare renders its challenge in a frame of its own, so read every
        // frame, not just the top document.
        let body = "";
        for (let i = 0; i < 8; i++) {
          body = await pageText(page);
          if (bodyNeedsVerification(body)) break;
          await page.waitForTimeout(1000);
        }
        if (bodyNeedsVerification(body)) {
          const verifyWaitMs = Number(process.env.M365_PERSONAL_VERIFY_WAIT_MS ?? 300_000);
          opts.onEvent?.(
            `Microsoft asks for human verification in the opened window — complete it there; waiting up to ${Math.round(verifyWaitMs / 1000)} s, then sending again`,
          );
          log.debug(`verification gate detected; waiting verifyWaitMs=${verifyWaitMs}`);
          const deadline = Date.now() + verifyWaitMs;
          let cleared = false;
          while (Date.now() < deadline) {
            const text = await pageText(page);
            if (!bodyNeedsVerification(text)) {
              cleared = true;
              break;
            }
            await page.waitForTimeout(1000);
          }
          if (cleared) {
            const composer2 = page
              .locator('#m365-chat-editor-target-element, textarea, [role="textbox"], [contenteditable="true"]')
              .first();
            await composer2.click({ timeout: 15_000 });
            await page.keyboard.press("Control+A");
            await page.keyboard.press("Backspace");
            await page.keyboard.type(PROMPT_TO_TRIGGER_SOCKET, { delay: 40 });
            const send2 = page.locator('button[aria-label*="Send" i], button[title*="Send" i]').first();
            if (await send2.count()) await send2.click({ timeout: 10_000 });
            else await page.keyboard.press("Enter");
            socket = await waitForSocket(waitMs);
          }
          if (!socket && !cleared) {
            throw new PersonalCaptureError(
              "no_chat_frame",
              "Microsoft asked for human verification (security check) and it was not completed in time",
            );
          }
        }
        if (!socket) throw new PersonalCaptureError("no_chat_frame");
      }
    }

    const capture = parsePersonalCapture(socket.url, socket.frames);
    log.debug(
      `captured scenario=${capture.scenario} expiresAt=${new Date(capture.expiresAt).toISOString()}`,
    );
    return capture;
  } finally {
    // An attached Chrome belongs to the operator: disconnect, never close it.
    if (attached) await attached.close().catch(() => {});
    else await context.close().catch((err) => log.debug("closing the capture browser failed:", err));
  }
}
