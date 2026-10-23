import { isPersonalMode } from "@m365-copilot/core";
import { ACCOUNTS_UI_HTML } from "../../ui/accounts-page";

/**
 * GET /accounts/ui — the accounts dashboard (§7.5). The page is
 * `ui/accounts.html`, inlined into this bundle at build time; it reads and
 * writes everything through the /accounts JSON routes from the browser, so this
 * handler serves the page and nothing else. 404 in enterprise mode, like the
 * rest of the accounts API.
 */
export default defineEventHandler((event) => {
  if (!isPersonalMode()) {
    setResponseStatus(event, 404);
    return { error: { code: "not_found", message: "account management requires personal auth mode" } };
  }
  setHeader(event, "content-type", "text/html; charset=utf-8");
  setHeader(event, "cache-control", "no-store");
  return ACCOUNTS_UI_HTML;
});
