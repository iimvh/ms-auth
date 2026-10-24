import * as msal from "@azure/msal-node";
import {
  readFileSync,
  writeFileSync,
  existsSync,
  mkdirSync,
} from "node:fs";
import { execSync } from "node:child_process";
import { join } from "node:path";
import { homedir } from "node:os";

const CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const AUTHORITY = "https://login.microsoftonline.com/common";
const REDIRECT_URI =
  "https://login.microsoftonline.com/common/oauth2/nativeclient";
const SCOPES = [
  "https://substrate.office.com/sydney/M365Chat.Read",
  "https://substrate.office.com/sydney/sydney.readwrite",
];

// Generated-image bytes live on designerapp.officeapps.live.com behind SharePoint
// Embedded (§14). The Sydney token 401s there; the artifact wants a token for the
// designerapp *service*. Our own first-party client IS preauthorized for it (the
// web client mints it via refresh_token), so acquireTokenSilent works — confirmed
// fetching a 2.3 MB PNG. Note it's an RSA-OAEP JWE, opaque to us; we pass it through.
const IMAGE_ARTIFACT_SCOPES = [
  "https://designerappservice.officeapps.live.com/.default",
];

export function getImageArtifactToken(): Promise<string | null> {
  return getTokenForScope(IMAGE_ARTIFACT_SCOPES);
}

import { createLogger } from "../../io/log.js";
import { z } from "zod/v4";
const log = createLogger("auth");

const CONFIG_DIR = join(homedir(), ".config", "opencode-m365");

function resolveFile(envVar: string, defaultName: string): string {
  if (process.env[envVar]) return process.env[envVar]!;
  mkdirSync(CONFIG_DIR, { recursive: true });
  return join(CONFIG_DIR, defaultName);
}

const CACHE_FILE = resolveFile("M365_CACHE_FILE", "msal-cache.json");
const SECRETS_FILE = resolveFile("M365_SECRETS_FILE", "secrets.json");

const LOGIN_LOCALE = process.env.M365_LOGIN_LOCALE ?? "en-GB";
const LOGIN_TIMEZONE = process.env.M365_LOGIN_TIMEZONE ?? "Europe/Copenhagen";

export function interactiveLoginPermitted(
  env: Record<string, string | undefined> = process.env,
): boolean {
  return env[INTERACTIVE_APPROVAL_ENV] === "1" && env[NO_INTERACTIVE_ENV] !== "1";
}

function interactiveAllowed(): boolean {
  return process.env[NO_INTERACTIVE_ENV] !== "1";
}

function interactiveTimeoutMs(): number {
  const raw = Number(process.env.M365_INTERACTIVE_TIMEOUT_MS ?? 600_000);
  return Number.isFinite(raw) && raw > 0 ? raw : 600_000;
}

function loadCache(app: msal.PublicClientApplication) {
  if (existsSync(CACHE_FILE)) {
    try {
      app.getTokenCache().deserialize(readFileSync(CACHE_FILE, "utf-8"));
    } catch {}
  }
}

function saveCache(app: msal.PublicClientApplication) {
  try {
    writeFileSync(CACHE_FILE, app.getTokenCache().serialize());
  } catch {}
}

let _app: msal.PublicClientApplication | null = null;

function getApp(): msal.PublicClientApplication {
  if (!_app) {
    _app = new msal.PublicClientApplication({
      auth: { clientId: CLIENT_ID, authority: AUTHORITY },
    });
    loadCache(_app);
  }
  return _app;
}

async function buildAuthUrlForScopes(app: msal.PublicClientApplication, scopes: string[]) {
  const cryptoProvider = new msal.CryptoProvider();
  const { verifier, challenge } = await cryptoProvider.generatePkceCodes();

  const authUrl = await app.getAuthCodeUrl({
    scopes,
    redirectUri: REDIRECT_URI,
    codeChallenge: challenge,
    codeChallengeMethod: "S256",
  });

  return { authUrl, verifier };
}

const LOGIN_DEBUG_DIR = join(CONFIG_DIR, "login-debug");

const BROWSER_PROFILE_DIR = resolveFile("M365_BROWSER_PROFILE", "browser-profile");

const LOGIN_USER_AGENT =
  process.env.M365_LOGIN_UA ??
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/146.0.0.0 Safari/537.36";

export interface Credentials {
  email: string;
  password: string;
  mfaSecret: string;
}

export const SecretsFileSchema = z.object({
  email: z.string().min(1),
  password: z.string().min(1),
  mfaSecret: z.string().min(1),
});

export const INTERACTIVE_APPROVAL_ENV = "M365_ENABLE_INTERACTIVE_APPROVAL";
export const NO_INTERACTIVE_ENV = "M365_NO_INTERACTIVE";

function resolveChromiumPath(): string | undefined {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;
  for (const bin of ["chromium", "chromium-browser", "google-chrome", "chrome"]) {
    try {
      const found = execSync(`command -v ${bin}`, { stdio: ["ignore", "pipe", "ignore"] })
        .toString()
        .trim();
      if (found) {
        log.info(`Resolved system browser: ${found}`);
        return found;
      }
    } catch {
    }
  }
  return undefined;
}

async function capture(page: any, label: string): Promise<void> {
  try {
    mkdirSync(LOGIN_DEBUG_DIR, { recursive: true });
    await page.screenshot({
      path: join(LOGIN_DEBUG_DIR, `${label}.png`),
      fullPage: true,
    });
    writeFileSync(join(LOGIN_DEBUG_DIR, `${label}.html`), await page.content());
    writeFileSync(join(LOGIN_DEBUG_DIR, `${label}.url.txt`), page.url());
    log.info(`Captured ${label} — url: ${page.url()}`);
  } catch (e: any) {
    log.error(`Failed to capture ${label}: ${e?.message}`);
  }
}

async function fillVerified(
  page: any,
  selector: string,
  value: string,
  label: string,
): Promise<void> {
  const loc = page.locator(`${selector}:visible`).first();
  await loc.waitFor({ state: "visible", timeout: 20000 });
  await loc.click();
  await loc.fill(value);
  let got = await loc.inputValue();
  if (got !== value) {
    log.info(`${label}: fill mismatch (${got.length} chars), retyping`);
    await loc.fill("");
    await loc.pressSequentially(value, { delay: 20 });
    got = await loc.inputValue();
  }
  if (got !== value) {
    throw new Error(`${label}: field still empty after refill`);
  }
}

async function clickSubmit(page: any): Promise<void> {
  await page
    .locator('input[type="submit"]:visible, button[type="submit"]:visible')
    .first()
    .click();
}

async function isVisibleSoon(page: any, selector: string, timeout: number): Promise<boolean> {
  try {
    await page.locator(`${selector}:visible`).first().waitFor({ state: "visible", timeout });
    return true;
  } catch {
    return false;
  }
}

async function clickAccountTileIfPresent(page: any, email: string): Promise<boolean> {
  try {
    await page.locator("#tilesHolder:visible").first().waitFor({ state: "visible", timeout: 5000 });
  } catch {
    return false;
  }
  const tile = page
    .locator(
      `[data-test-id="${email.toLowerCase()}"]:visible, ` +
      `#tilesHolder .tile [role="button"][data-test-id]:not([data-test-id$="-menu-dots"]):visible`,
    )
    .first();
  try {
    await tile.waitFor({ state: "visible", timeout: 5000 });
    await tile.click();
    log.info("Account picker — clicked remembered account tile (SSO)");
    return true;
  } catch {
    return false;
  }
}

async function driveAzureLogin(page: any, creds: Credentials): Promise<void> {
  const { TOTP } = await import("otpauth");

  await capture(page, "step0-landing");

  const picked = await clickAccountTileIfPresent(page, creds.email);

  if (!picked && (await isVisibleSoon(page, 'input[name="loginfmt"]', 8000))) {
    log.info("Step: email");
    await fillVerified(page, 'input[name="loginfmt"]', creds.email, "email");
    await clickSubmit(page);
    await capture(page, "step1-after-email");
  } else {
    log.info(`Step: email skipped (${picked ? "picked account tile" : "SSO — no email prompt"})`);
  }

  if (await isVisibleSoon(page, 'input[name="passwd"]', 8000)) {
    log.info("Step: password");
    await fillVerified(page, 'input[name="passwd"]', creds.password, "password");
    await clickSubmit(page);
    await capture(page, "step2-after-password");
  } else {
    log.info("Step: password skipped (SSO — no password prompt)");
  }

  if (await isVisibleSoon(page, 'input[name="otc"]', 8000)) {
    log.info("Step: mfa");
    const otpCode = new TOTP({ secret: creds.mfaSecret }).generate();
    await fillVerified(page, 'input[name="otc"]', otpCode, "otc");
    await clickSubmit(page);
    await capture(page, "step3-after-mfa");
  } else {
    log.info("Step: mfa skipped (SSO — no TOTP prompt)");
  }

  log.info("Step: stay-signed-in");
  try {
    await page.locator("#idSIButton9:visible").click({ timeout: 8000 });
  } catch {
  }
}

async function runBrowserLogin(
  app: msal.PublicClientApplication,
  scopes: string[],
  creds: Credentials,
  attempts = 3,
): Promise<string | null> {
  const { chromium } = await import("playwright");

  for (let attempt = 1; attempt <= attempts; attempt++) {
    const { authUrl, verifier } = await buildAuthUrlForScopes(app, scopes);
    const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
      headless: true,
      executablePath: resolveChromiumPath(),
      args: [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-blink-features=AutomationControlled",
      ],
      userAgent: LOGIN_USER_AGENT,
      locale: LOGIN_LOCALE,
      timezoneId: LOGIN_TIMEZONE,
      viewport: { width: 1280, height: 800 },
    });
    await context.addInitScript(() => {
      Object.defineProperty(navigator, "webdriver", { get: () => undefined });
    });
    const page = context.pages()[0] ?? (await context.newPage());

    let resolveCode: (code: string) => void;
    const codePromise = new Promise<string>((res) => {
      resolveCode = res;
    });
    page.on("request", (req: any) => {
      const u = req.url();
      if (u.includes("/oauth2/nativeclient") && u.includes("code=")) {
        const c = new URL(u).searchParams.get("code");
        if (c) {
          log.info("Captured auth code from nativeclient redirect");
          resolveCode(c);
        }
      }
    });

    try {
      log.info(`Browser login attempt ${attempt}/${attempts} for [${scopes.join(", ")}]`);
      await page.goto(authUrl, { waitUntil: "domcontentloaded" });
      const drive = driveAzureLogin(page, creds).catch((e: any) =>
        log.info(`driveAzureLogin ended early: ${e?.message}`),
      );

      let timer: ReturnType<typeof setTimeout> | undefined;
      const authCode = await Promise.race([
        codePromise,
        new Promise<string>((_, rej) => {
          timer = setTimeout(() => rej(new Error("Timed out waiting for auth code")), 45000);
        }),
      ]).finally(() => clearTimeout(timer));
      void drive;

      const result = await app.acquireTokenByCode({
        code: authCode,
        scopes,
        redirectUri: REDIRECT_URI,
        codeVerifier: verifier,
      });
      saveCache(app);
      log.info(`Browser login succeeded as ${result.account?.username}`);
      return result.accessToken;
    } catch (err: any) {
      await capture(page, `attempt-${attempt}-fail`);
      log.error(`Browser login attempt ${attempt}/${attempts} failed: ${err.message}`);
      if (attempt < attempts) {
        await new Promise((r) => setTimeout(r, 31_000));
      }
    } finally {
      await context.close();
    }
  }
  return null;
}

export async function getTokenSilent(): Promise<string | null> {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  if (accounts.length === 0) return null;

  try {
    const result = await app.acquireTokenSilent({
      scopes: SCOPES,
      account: accounts[0],
    });
    saveCache(app);
    return result.accessToken;
  } catch {
    return null;
  }
}

export async function driveSignInForm(
  page: unknown,
  creds: { email: string; password: string; mfaSecret: string },
): Promise<void> {
  await driveAzureLogin(page, creds);
}

export async function loginAutomated(
  email: string,
  password: string,
  mfaSecret: string,
): Promise<string> {  const app = getApp();
  log.info("Starting automated login...");
  const token = await runBrowserLogin(
    app,
    SCOPES,
    { email, password, mfaSecret },
    1,
  );
  if (!token) {
    throw new Error(
      `Automated login failed — see artifacts in ${LOGIN_DEBUG_DIR}`,
    );
  }
  return token;
}

async function loginInteractiveForScopes(scopes: string[]): Promise<string> {
  if (!interactiveAllowed()) {
    throw new Error("Interactive login is disabled (M365_NO_INTERACTIVE=1)");
  }
  const { chromium } = await import("playwright");
  const app = getApp();
  const timeoutMs = interactiveTimeoutMs();
  const { authUrl, verifier } = await buildAuthUrlForScopes(app, scopes);

  const context = await chromium.launchPersistentContext(BROWSER_PROFILE_DIR, {
    headless: false,
    executablePath: resolveChromiumPath(),
    args: [
      "--no-sandbox",
      "--disable-dev-shm-usage",
      "--disable-blink-features=AutomationControlled",
    ],
    userAgent: LOGIN_USER_AGENT,
    locale: LOGIN_LOCALE,
    timezoneId: LOGIN_TIMEZONE,
    viewport: { width: 1280, height: 800 },
  });
  await context.addInitScript(() => {
    Object.defineProperty(navigator, "webdriver", { get: () => undefined });
  });
  const page = context.pages()[0] ?? (await context.newPage());

  let resolveCode: (code: string) => void;
  const codePromise = new Promise<string>((res) => {
    resolveCode = res;
  });
  page.on("request", (req: any) => {
    const u = req.url();
    if (u.includes("/oauth2/nativeclient") && u.includes("code=")) {
      const c = new URL(u).searchParams.get("code");
      if (c) {
        log.info("Captured auth code from nativeclient redirect (interactive)");
        resolveCode(c);
      }
    }
  });

  try {
    console.error("[m365 auth] Interactive approval required.");
    console.error("[m365 auth] A browser window has opened — complete sign-in there.");
    console.error(`[m365 auth] Waiting up to ${Math.round(timeoutMs / 1000)}s.`);
    await page.goto(authUrl, { waitUntil: "domcontentloaded" });
    let timer: ReturnType<typeof setTimeout> | undefined;
    const authCode = await Promise.race([
      codePromise,
      new Promise<string>((_, rej) => {
        timer = setTimeout(
          () => rej(new Error(`Timed out waiting for interactive auth code after ${timeoutMs}ms`)),
          timeoutMs,
        );
      }),
    ]).finally(() => clearTimeout(timer));

    const result = await app.acquireTokenByCode({
      code: authCode,
      scopes,
      redirectUri: REDIRECT_URI,
      codeVerifier: verifier,
    });
    saveCache(app);
    log.info(`Interactive login succeeded as ${result.account?.username}`);
    return result.accessToken;
  } catch (err: any) {
    await capture(page, "interactive-fail");
    throw err;
  } finally {
    await context.close();
  }
}

let inflightReauth: Promise<boolean> | null = null;

export function forceReauth(): Promise<boolean> {
  return (inflightReauth ??= doForceReauth().finally(() => {
    inflightReauth = null;
  }));
}

async function doForceReauth(): Promise<boolean> {
  try {
    const silent = await getTokenSilent();
    if (silent) {
      log.info("forceReauth: refreshed silently — no interactive login needed");
      return true;
    }
    const secrets = loadSecrets();
    const canPromptHuman = interactiveLoginPermitted();
    if (!secrets) {
      if (canPromptHuman) {
        log.info("forceReauth: no secrets file, asking for interactive approval");
        await loginInteractiveForScopes(SCOPES);
        return true;
      }
      log.error("forceReauth: silent refresh failed and no secrets file — cannot re-login");
      return false;
    }
    log.info("forceReauth: silent unavailable, doing automated login");
    try {
      await loginAutomated(secrets.email, secrets.password, secrets.mfaSecret);
      log.info("forceReauth: automated login succeeded");
      return true;
    } catch (err: any) {
      if (!canPromptHuman) throw err;
      log.info(`forceReauth: automated login failed (${err.message}), asking for interactive approval`);
      await loginInteractiveForScopes(SCOPES);
      return true;
    }
  } catch (err: any) {
    log.error(`forceReauth failed: ${err.message}`);
    return false;
  }
}

export function loadSecrets(): Credentials | null {
  if (!existsSync(SECRETS_FILE)) return null;
  try {
    const parsed = SecretsFileSchema.safeParse(JSON.parse(readFileSync(SECRETS_FILE, "utf-8")));
    if (parsed.success) return parsed.data;
  } catch {}
  return null;
}

export async function getTokenForScope(scopes: string[]): Promise<string | null> {
  const app = getApp();
  const accounts = await app.getTokenCache().getAllAccounts();
  log.info(`getTokenForScope: ${scopes.join(",")} — ${accounts.length} accounts in cache`);

  if (accounts.length > 0) {
    try {
      const result = await app.acquireTokenSilent({
        scopes,
        account: accounts[0],
      });
      saveCache(app);
      return result.accessToken;
    } catch (err: any) {
      log.info(`getTokenForScope: silent failed (${err.message}), trying browser login`);
    }
  }

  const secrets = loadSecrets();
  const canPromptHuman = interactiveLoginPermitted();
  if (!secrets) return canPromptHuman ? loginInteractiveForScopes(scopes) : null;
  try {
    return await runBrowserLogin(app, scopes, secrets);
  } catch (err: any) {
    if (!canPromptHuman) throw err;
    log.info(`getTokenForScope: browser login failed (${err.message}), asking for interactive approval`);
    return loginInteractiveForScopes(scopes);
  }
}

let inflightToken: Promise<string> | null = null;

export function getToken(): Promise<string> {
  return (inflightToken ??= doGetToken().finally(() => {
    inflightToken = null;
  }));
}

async function doGetToken(): Promise<string> {
  const silent = await getTokenSilent();
  if (silent) {
    log.info("Token refreshed silently");
    return silent;
  }

  const secrets = loadSecrets();
  const canPromptHuman = interactiveLoginPermitted();
  if (!secrets) {
    if (canPromptHuman) return loginInteractiveForScopes(SCOPES);
    throw new Error(
      "No cached token and no secrets.json — cannot authenticate. Provide email/password/mfaSecret for automated login, or set M365_ENABLE_INTERACTIVE_APPROVAL=1 to sign in by hand in a visible browser (needed for tenants without TOTP: push-only MFA, FIDO2, Okta/Ping/Duo).",
    );
  }
  if (!canPromptHuman) {
    return loginAutomated(secrets.email, secrets.password, secrets.mfaSecret);
  }
  try {
    return await loginAutomated(secrets.email, secrets.password, secrets.mfaSecret);
  } catch (err: any) {
    log.info(`Automated login failed (${err.message}), asking for interactive approval`);
    return loginInteractiveForScopes(SCOPES);
  }
}
