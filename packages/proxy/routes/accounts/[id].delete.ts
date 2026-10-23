import { accountsDeleteResponse } from "@m365-copilot/proxy-lib";
import { accountsDeps } from "../../server-accounts";

/**
 * DELETE /accounts/:id — removes the entry (its files move to .trash/).
 * Refuses while that account is serving a conversation (409 account_busy).
 */
export default defineEventHandler(async (event) => {
  const id = getRouterParam(event, "id") ?? "";
  const { result } = accountsDeleteResponse(id, accountsDeps);
  setResponseStatus(event, result.status);
  return result.body;
});
