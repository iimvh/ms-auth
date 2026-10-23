import { accountsListResponse, type AccountsDeps } from "@m365-copilot/proxy-lib";
import { pool } from "./server-pool";

/**
 * The real deps for the account routes: the process-wide pool (so a delete can
 * see whether that account is serving a thread) and the default CDP probe
 * (each account's own port, taken from the registry).
 */
export const accountsDeps: AccountsDeps = {
  pool: {
    liveAccountIds: () => pool.liveAccountIds(),
    forgetAccount: (id) => pool.forgetAccount(id),
  },
};
