import { accountsRenameResponse } from "@m365-copilot/proxy-lib";
import { accountsDeps } from "../../server-accounts";

/** PATCH /accounts/:id { label } — rename only; the id is the identity. */
export default defineEventHandler(async (event) => {
  const id = getRouterParam(event, "id") ?? "";
  const input = (await readBody(event)) ?? {};
  const r = accountsRenameResponse(id, { label: input.label }, accountsDeps);
  setResponseStatus(event, r.status);
  return r.body;
});
