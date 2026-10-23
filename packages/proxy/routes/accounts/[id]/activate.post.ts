import { accountsActivateResponse } from "@m365-copilot/proxy-lib";
import { accountsDeps } from "../../../server-accounts";

/**
 * POST /accounts/:id/activate — the switch. It takes effect on the next request
 * (the account is resolved per request, §7.3); a thread already running keeps
 * the account it was opened on.
 */
export default defineEventHandler(async (event) => {
  const id = getRouterParam(event, "id") ?? "";
  const r = accountsActivateResponse(id, accountsDeps);
  setResponseStatus(event, r.status);
  return r.body;
});
