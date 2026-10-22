import { getToken, resolveAuthMode } from "@m365-copilot/core";

/**
 * Authenticate against M365 once, at server startup. A failure here throws and
 * aborts boot — the equivalent of the old binary's `process.exit(1)` on auth
 * failure, so the server never comes up half-broken.
 */
export default defineNitroPlugin(async () => {
  console.log("Authenticating...");
  try {
    // Resolves first so an unknown M365_AUTH value fails at boot naming the
    // allowed values, before any token work happens.
    const mode = resolveAuthMode();
    console.log(`M365_AUTH=${mode}`);
    await getToken();
  } catch (err: any) {
    console.error(`Auth failed: ${err.message}`);
    throw err;
  }
});
