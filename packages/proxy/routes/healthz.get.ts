import { buildHealthzPayload } from "@m365-copilot/proxy-lib";

/**
 * Credential health. `/health` answers "is the process up?"; this answers "can
 * this process still serve a turn?", which in personal mode is the thing that
 * fails silently behind a generic 502 (issue.md, candidate fix 3).
 *
 * It never mints and never spends a turn: it reads the ACTIVE account's record
 * and probes that account's CDP endpoint. Always 200 — the body says how bad it
 * is, so a gateway can tell "upstream reachable but credential stale" from
 * "proxy down". `accounts` carries the count and the active id (§7.8), so a
 * switch can be confirmed from here without a restart.
 */
export default defineEventHandler(async () => {
  return buildHealthzPayload();
});
