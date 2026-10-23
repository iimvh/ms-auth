import { accountsAddResponse } from "@m365-copilot/proxy-lib";
import { accountsDeps } from "../../server-accounts";

/**
 * POST /accounts { id, label } → 201 with the Chrome command to start that
 * account's browser and the login command to capture its session.
 */
export default defineEventHandler(async (event) => {
  const input = (await readBody(event)) ?? {};
  const r = await accountsAddResponse({ id: input.id, label: input.label }, accountsDeps);
  setResponseStatus(event, r.status);
  return r.body;
});
