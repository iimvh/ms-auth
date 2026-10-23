import { createLogger } from "../../log.js";
import type { Locator, Page } from "playwright";
import { resolveAccount } from "../../account-registry.js";
import { HarvestError, type HarvestedPage, type HarvestedSocket } from "./personal-harvest.js";
import { frameText } from "./personal-capture.js";

const log = createLogger("personal-harvest-cdp");

/** How long the composer may stay disabled before the harvest gives up. Copilot
 *  streams an answer for a while; a mint that lands mid-stream has to wait it
 *  out, because nothing it types would be sent. */
const COMPOSER_READY_MS = Number(process.env.M365_PERSONAL_COMPOSER_READY_MS ?? 45_000);

/** The CDP endpoint an operator's Chrome listens on — the active account's,
 *  so a mint can never attach to another account's Chrome (§7.2). */
export function personalCdpUrl(): string {
  return resolveAccount().cdpUrl;
}

/** The locale-tolerant locator for the chat composer (`capturePersonalSession`
 *  uses the same selector; the editor is a contenteditable span). */
const COMPOSER_SELECTOR =
  '#m365-chat-editor-target-element, textarea, [role="textbox"], [contenteditable="true"]';

/** A URL that is the AAD sign-in flow rather than the chat app. */
function isLoginUrl(url: string): boolean {
  return url.includes("login.microsoftonline.com") || url.includes("/login");
}

function isChathubSocket(url: string): boolean {
  try {
    const p = new URL(url);
    return p.host === "substrate.office.com" && p.pathname.startsWith("/m365Copilot/Chathub/");
  } catch {
    return false;
  }
}

/** Whether a payload is the chat invocation (type:4, target:"chat"). */
function isChatInvocation(text: string): boolean {
  return text.includes('"target":"chat"') || text.includes('"target": "chat"');
}

/** Reach the operator's Chrome over CDP and hand back its chat page.
 *
 *  This is deliberately NOT a launch: a Playwright-launched browser is
 *  fingerprinted and cannot pass Cloudflare (F101), and on Windows the consumer
 *  SSO does not auto-continue in a cold headless browser — which is exactly why
 *  the old re-mint failed while a readable error was never produced (issue.md).
 *  Everything here connects to something already running. */
export async function attachChatPage(): Promise<HarvestedPage> {
  const cdpUrl = personalCdpUrl();
  const { chromium } = await import("playwright");
  let browser;
  try {
    browser = await chromium.connectOverCDP(cdpUrl, { timeout: 10_000 });
  } catch (err) {
    throw new HarvestError("chrome_unreachable", `could not connect over CDP to ${cdpUrl}: ${message(err)}`);
  }
  log.debug(`attached over CDP to ${cdpUrl}`);

  const context = browser.contexts()[0] ?? (await browser.newContext());
  const page = context.pages().find((p) => p.url().includes("m365.cloud.microsoft"));
  if (!page) {
    await browser.close().catch(() => {});
    throw new HarvestError("no_chat_tab", "the attached Chrome has no m365.cloud.microsoft page open");
  }
  if (isLoginUrl(page.url())) {
    await browser.close().catch(() => {});
    throw new HarvestError("session_dead", `the chat tab is on a sign-in page: ${page.url()}`);
  }

  // Sockets the page opens are only observable from now on, so the listener is
  // installed before the composer is touched.
  const sockets: HarvestedSocket[] = [];
  page.on("websocket", (ws) => {
    const rec: HarvestedSocket = { url: ws.url(), frames: [] };
    sockets.push(rec);
    ws.on("framesent", (payload) => {
      const text = frameText(payload);
      if (text) rec.frames.push(text);
    });
  });

  const close = async () => {
    // The Chrome belongs to the operator: disconnect, never close it.
    await browser!.close().catch((err: unknown) => log.debug("disconnecting from the operator's Chrome failed:", err));
  };

  return {
    url: () => page.url(),
    async sendAndWait(waitMs: number): Promise<HarvestedSocket | null> {
      await page.bringToFront().catch(() => {});
      const composer = page.locator(COMPOSER_SELECTOR).first();

      // The composer is DISABLED while Copilot is still answering the previous
      // probe, and a click on a disabled editor does nothing: the message is
      // never sent, no socket opens, and the harvest used to sit out its full
      // timeout and report `mint_not_rotated`. Measured both sides — a clean
      // page is `enabled: true` and opens its socket within ~1 s; a page still
      // streaming is `enabled: false` and never sends. Wait for interactivity
      // before typing, so "no socket" means the page really refused.
      const ready = await waitForComposer(composer, page, COMPOSER_READY_MS);
      if (!ready) {
        throw new HarvestError(
          "capture_failed",
          `the chat composer stayed disabled for ${COMPOSER_READY_MS} ms (the page is busy or blocked); no message was sent`,
        );
      }

      await composer.click({ timeout: 30_000 });
      // The page restores unsent drafts, so clear before typing (F101).
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(process.env.M365_PERSONAL_MINT_PROMPT ?? ".", { delay: 40 });
      const send = page.locator('button[aria-label*="Send" i], button[title*="Send" i]').first();
      if (await send.count()) await send.click({ timeout: 10_000 });
      else await page.keyboard.press("Enter");

      const deadline = Date.now() + waitMs;
      while (Date.now() < deadline) {
        const found = sockets.find((s) => isChathubSocket(s.url) && s.frames.some(isChatInvocation));
        if (found) return found;
        await page.waitForTimeout(250);
      }
      log.debug(`no Chathub socket with a chat invocation frame within ${waitMs} ms`);
      return null;
    },
    close,
  };
}

/** Wait until the composer accepts input. Returns false if it never does. */
async function waitForComposer(
  composer: Locator,
  page: Page,
  timeoutMs: number,
): Promise<boolean> {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (await composer.isEnabled().catch(() => false)) return true;
    await page.waitForTimeout(250);
  }
  return false;
}

function message(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}
