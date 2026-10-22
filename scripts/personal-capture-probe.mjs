#!/usr/bin/env node
/**
 * RE probe for the personal (consumer) capture: where does the chat invocation
 * actually travel? Opens the real chat page in the capture profile, signs in with
 * the stored credentials, logs EVERY websocket and EVERY sent frame, then types a
 * message through the page's composer and logs what appears.
 *
 * Run from packages/core (playwright resolves there) with the repo's dist:
 *   M365_DEBUG=1 node ../../scripts/personal-capture-probe.mjs
 */
import { mkdirSync } from "node:fs";
import { execSync } from "node:child_process";
import { chromium } from "playwright";
import {
  loadSecrets,
  driveSignInForm,
  personalProfileDir,
  isLoginUrl,
} from "../packages/core/dist/index.mjs";

const OUT = process.env.M365_FRAME_DIR ?? "C:/Users/hac/AppData/Local/Temp/opencode/personal-probe";
mkdirSync(OUT, { recursive: true });

function findChromium() {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;
  for (const bin of ["chromium", "google-chrome", "chrome"]) {
    try {
      return execSync(`command -v ${bin}`, { stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
    } catch {}
  }
  return undefined;
}

const ctx = await chromium.launchPersistentContext(personalProfileDir(), {
  headless: false,
  executablePath: findChromium(),
  args: ["--no-sandbox", "--disable-dev-shm-usage", "--disable-blink-features=AutomationControlled"],
  userAgent: process.env.M365_LOGIN_UA ??
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36",
  locale: "en-GB",
  viewport: { width: 1280, height: 800 },
});
const page = ctx.pages()[0] ?? (await ctx.newPage());
const sockets = new Map();
page.on("websocket", (ws) => {
  const rec = { url: ws.url(), frames: [] };
  sockets.set(rec, ws);
  console.log(`[ws open] ${ws.url().slice(0, 90)}`);
  ws.on("framesent", (p) => {
    const d = p && typeof p === "object" && "payload" in p ? p.payload : p;
    const text = typeof d === "string" ? d : Buffer.from(d).toString("utf8");
    rec.frames.push(text);
    console.log(`[frame ${ws.url().includes("chathub") || ws.url().includes("Chathub") ? "CHATHUB" : ws.url().includes("trouter") ? "trouter" : "other"}] ${text.slice(0, 220).replace(/\n/g, " ")}`);
  });
});

await page.goto("https://m365.cloud.microsoft/chat", { waitUntil: "domcontentloaded" });
console.log("[landing]", page.url().slice(0, 100));
if (isLoginUrl(page.url())) {
  const creds = loadSecrets();
  await driveSignInForm(page, creds).catch((e) => console.log("[signin]", e?.message));
  for (let i = 0; i < 60 && isLoginUrl(page.url()); i++) await page.waitForTimeout(500);
}
console.log("[after sign-in]", page.url().slice(0, 100));
await page.waitForTimeout(12000);
await page.screenshot({ path: `${OUT}/chat-loaded.png` });

const candidates = await page.evaluate(() => {
  const pick = (sel) =>
    Array.from(document.querySelectorAll(sel)).slice(0, 6).map((el) => ({
      tag: el.tagName,
      id: el.id,
      role: el.getAttribute("role"),
      name: el.getAttribute("name"),
      ph: el.getAttribute("placeholder") ?? el.getAttribute("aria-label") ?? el.getAttribute("data-placeholder"),
      ce: el.getAttribute("contenteditable"),
      vis: !!(el.offsetParent || el.getClientRects().length),
    }));
  return {
    textarea: pick("textarea"),
    textboxRoles: pick('[role="textbox"]'),
    contentEditable: pick('[contenteditable="true"]'),
    named: pick('[data-testid*="composer"], [id*="composer"], [class*="composer"]'),
  };
});
console.log("[composer candidates]", JSON.stringify(candidates, null, 1).slice(0, 2500));

const box = page.locator('[contenteditable="true"], textarea, [role="textbox"]').first();
try {
  await box.click({ timeout: 5000 });
  await box.fill("hi").catch(async () => { await box.type("hi"); });
  await page.waitForTimeout(500);
  await page.screenshot({ path: `${OUT}/typed.png` });
  await box.press("Enter");
  console.log("[typed] sent 'hi'");
} catch (e) {
  console.log("[typed] failed:", e?.message);
}
await page.waitForTimeout(25000);
await page.screenshot({ path: `${OUT}/after-send.png` });

const summary = [...sockets.keys()].map((r) => ({ url: r.url.slice(0, 80), frames: r.frames.length }));
console.log("[summary]", JSON.stringify(summary));
await ctx.close();
process.exit(0);
