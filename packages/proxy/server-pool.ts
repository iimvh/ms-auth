import { SessionPool } from "@m365-copilot/proxy-lib";

/**
 * Process-wide session pool shared by every request.
 *
 * The pool maps each distinct conversation to its own M365 session, so a single
 * pool for the whole server is exactly the behaviour the old `createApp()` had
 * (it created one pool per app instance).
 */
export const pool = new SessionPool();
