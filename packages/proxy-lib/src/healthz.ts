import { isPersonalMode, personalHealth, registryStatus, type PersonalHealth } from "@m365-copilot/core";

/**
 * The `/healthz` payload (spec §7.8). Personal mode reports the ACTIVE account's
 * credential state and how many accounts the registry holds; enterprise keeps
 * the shape it has always had, with `accounts` added and `credential` marked
 * not applicable.
 *
 * It never throws: a broken registry must show up as a report, not a 500 on the
 * one endpoint an operator uses to find out what is wrong.
 */
export interface HealthzPayload {
  status: "ok";
  authMode: "msal" | "personal";
  credential: PersonalHealth | { applicable: false };
  accounts: {
    count: number;
    activeId: string | null;
    /** `registry` when accounts.json decided it, `builtin` when it fell back. */
    repository: "registry" | "builtin";
    problem?: { code: string; detail: string; hint: string };
  };
}

export async function buildHealthzPayload(): Promise<HealthzPayload> {
  const accounts = safeRegistryStatus();
  if (!isPersonalMode()) {
    return { status: "ok", authMode: "msal", credential: { applicable: false }, accounts };
  }
  return { status: "ok", authMode: "personal", credential: await personalHealth(), accounts };
}

/** The registry's view, degrading to a readable report instead of throwing. */
function safeRegistryStatus(): HealthzPayload["accounts"] {
  try {
    const st = registryStatus();
    return {
      count: st.count,
      activeId: st.activeId,
      repository: st.repository,
      ...(st.problem ? { problem: { code: st.problem.code, detail: st.problem.detail, hint: st.problem.hint } } : {}),
    };
  } catch (err) {
    return {
      count: 0,
      activeId: null,
      repository: "builtin",
      problem: { code: "unreadable", detail: err instanceof Error ? err.message : String(err), hint: "check accounts.json" },
    };
  }
}
