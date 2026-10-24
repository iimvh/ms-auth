import { createLogger } from "../../io/log.js";
import type { Locator, Page } from "playwright";
import { resolveAccount } from "../../account-registry.js";
import { HarvestError, type HarvestedPage, type HarvestedSocket } from "./personal-harvest.js";
import { frameText } from "./personal-capture.js";

const log = createLogger("personal-harvest-cdp");

const COMPOSER_READY_MS = Number(process.env.M365_PERSONAL_COMPOSER_READY_MS ?? 45_000);

export function personalCdpUrl(): string {
  return resolveAccount().cdpUrl;
}

const COMPOSER_SELECTOR =
  '#m365-chat-editor-target-element, textarea, [role="textbox"], [contenteditable="true"]';

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

function isChatInvocation(text: string): boolean {
  return text.includes('"target":"chat"') || text.includes('"target": "chat"');
}

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
    await browser!.close().catch((err: unknown) => log.debug("disconnecting from the operator's Chrome failed:", err));
  };

  return {
    url: () => page.url(),
    async sendAndWait(waitMs: number): Promise<HarvestedSocket | null> {
      await page.bringToFront().catch(() => {});
      const composer = page.locator(COMPOSER_SELECTOR).first();

      const ready = await waitForComposer(composer, page, COMPOSER_READY_MS);
      if (!ready) {
        throw new HarvestError(
          "capture_failed",
          `the chat composer stayed disabled for ${COMPOSER_READY_MS} ms (the page is busy or blocked); no message was sent`,
        );
      }

      await composer.click({ timeout: 30_000 });
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
