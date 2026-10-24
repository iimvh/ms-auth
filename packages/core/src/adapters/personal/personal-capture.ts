import { execSync } from "node:child_process";
import { resolveAccount } from "../../account-registry.js";
import { createLogger } from "../../io/log.js";
import type { PersonalCapture } from "./personal-session.js";

const log = createLogger("personal-capture");

export interface SignInCredentials {
  email: string;
  password: string;
  mfaSecret: string;
}

export interface CaptureDeps {
  loadSecrets: () => SignInCredentials | null;
  driveSignInForm: (page: unknown, creds: SignInCredentials) => Promise<void>;
}

const hostDeps: { current: CaptureDeps | null } = { current: null };

export function providePersonalCaptureDeps(deps: CaptureDeps): void {
  hostDeps.current = deps;
}

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
    super(reason === "missing_param" ? what.missing_param : detail ? `${what[reason]}: ${detail}` : what[reason]);
    this.name = "PersonalCaptureError";
  }
}

const CHATHUB_HOST = "substrate.office.com";
const CHATHUB_PATH_PREFIX = "/m365Copilot/Chathub/";

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

const STRING_PARAMS: readonly string[] = [
  "scenario",
  "variants",
  "licenseType",
  "agentHost",
];
const OPTIONAL_STRING_PARAMS: readonly string[] = ["product"];
const JSON_PARAMS: readonly string[] = ["source"];

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

async function pageText(page: {
  locator(sel: string): { innerText(): Promise<string> };
  frames(): Array<{ locator(sel: string): { innerText(): Promise<string> } }>;
}): Promise<string> {
  const parts: string[] = [];
  for (const frame of page.frames()) {
    try {
      parts.push(await frame.locator("body").innerText());
    } catch {
    }
  }
  return parts.join(String.fromCharCode(10));
}

export function bodyNeedsVerification(text: string): boolean {
  return /verification required to continue|verify you are human|security check required|couldn.t be completed\. refresh to try again/i.test(text);
}

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

export function splitMessages(rawFrame: string): unknown[] {
  const messages: unknown[] = [];
  for (const chunk of rawFrame.split(RS)) {
    for (const piece of splitConcatenatedJson(chunk)) {
      try {
        messages.push(JSON.parse(piece));
      } catch {
      }
    }
  }
  return messages;
}

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
    }
  }
  return undefined;
}

const LOGIN_USER_AGENT =
  process.env.M365_LOGIN_UA ??
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36";

const CHAT_HOME = "https://m365.cloud.microsoft/chat";

export function isLoginUrl(url: string): boolean {
  return url.includes("login.microsoftonline.com") || url.includes("/login");
}

async function storedSecrets(deps: CaptureDeps): Promise<SignInCredentials | null> {
  if (process.env.M365_PERSONAL_AUTOMATED_LOGIN === "0") return null;
  try {
    return deps.loadSecrets();
  } catch {
    return null;
  }
}

export function frameText(data: unknown): string | null {
  try {
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

export async function capturePersonalSession(opts: {
  headed: boolean;
  profileDir: string;
  waitMs?: number;
  onEvent?: (line: string) => void;
  deps?: CaptureDeps;
}): Promise<PersonalCapture> {  const waitMs = opts.waitMs ?? Number(process.env.M365_PERSONAL_MINT_WAIT_MS ?? 10_000);
  log.debug(`capturing personal session (${opts.headed ? "headed" : "headless"}), waitMs=${waitMs}`);
  const { chromium } = await import("playwright");
  // Cloudflare's "verify you are human" gate refuses a Playwright-launched
  // browser no matter how the window is dressed: the launch itself is the
  // fingerprint. With a registry (an active account) or M365_PERSONAL_CDP_URL
  // set, the operator starts a REAL Chrome (`--remote-debugging-port=<port>
  // --user-data-dir=<profile dir>`) and we attach over CDP instead — nothing to
  // launch, no automation flag (F101).
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
    if (isLoginUrl(page.url()) && opts.headed) {
      const loginWaitMs = Number(process.env.M365_PERSONAL_LOGIN_WAIT_MS ?? 300_000);
      opts.onEvent?.(
        `the browser opened on the Microsoft sign-in page — complete the sign-in there; waiting up to ${Math.round(loginWaitMs / 1000)} s`,
      );
      log.debug(`headed login waits for interactive sign-in, loginWaitMs=${loginWaitMs}`);
      const deps = opts.deps ?? hostDeps.current;
      const creds = deps ? await storedSecrets(deps) : null;
      if (creds && deps) {
        opts.onEvent?.("trying the stored credentials' automated sign-in");
        void deps.driveSignInForm(page, creds)
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
        const socket = sockets.find((s) => isChathub(s.url) && hasChatFrame(s.frames));
        if (socket) return socket;
        await page.waitForTimeout(250);
      }
      return null;
    };

    let socket = await waitForSocket(waitMs);
    if (!socket) {
      log.debug(`no chat invocation frame after ${waitMs} ms; typing a short message to force one`);
      const composer = page
        .locator('#m365-chat-editor-target-element, textarea, [role="textbox"], [contenteditable="true"]')
        .first();
      await composer.click({ timeout: 30_000 });
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(PROMPT_TO_TRIGGER_SOCKET, { delay: 40 });
      const send = page.locator('button[aria-label*="Send" i], button[title*="Send" i]').first();
      if (await send.count()) await send.click({ timeout: 10_000 });
      else await page.keyboard.press("Enter");
      socket = await waitForSocket(waitMs);
      if (!socket) {
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
    if (attached) await attached.close().catch(() => {});
    else await context.close().catch((err) => log.debug("closing the capture browser failed:", err));
  }
}
