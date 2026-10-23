import { accountsListResponse } from "@m365-copilot/proxy-lib";
import { accountsDeps } from "../../server-accounts";

/** GET /accounts — every account, its health, and where the registry lives. */
export default defineEventHandler(async (event) => {
  const r = await accountsListResponse(accountsDeps);
  setResponseStatus(event, r.status);
  return r.body;
});
