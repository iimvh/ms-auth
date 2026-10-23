import { getToken, resolveAuthMode } from "@m365-copilot/core";

/**
 * Authenticate against M365 once, at server startup.
 *
 * Enterprise (MSAL) mode keeps the old behaviour: a failure here throws and
 * aborts boot — the equivalent of the old binary's `process.exit(1)` — so the
 * server never comes up half-broken.
 *
 * Personal mode REPORTS instead (§7.8). The account is resolved per request, so
 * a dead record or a Chrome that is not running is a state the proxy recovers
 * from on the next request (start that account's Chrome, or switch to an
 * account that is alive). Aborting boot would take the accounts API and the UI
 * down with it — exactly what the operator needs to fix it.
 */
export default defineNitroPlugin(async () => {
  const mode = resolveAuthMode();
  console.log(`M365_AUTH=${mode}`);
  try {
    await getToken();
  } catch (err: any) {
    console.error(`Auth failed: ${err?.message ?? String(err)}`);
    if (mode !== "personal") throw err;
    console.error(
      "Booting anyway: personal mode resolves the account per request, so starting that account's Chrome (or switching accounts) recovers without a restart. See GET /healthz and GET /accounts/ui.",
    );
  }
});
