import { isPersonalMode, personalHealth, registryStatus, type PersonalHealth } from "@m365-copilot/core";

// /healthz must not mint or spend a turn; a design constraint.
export interface HealthzPayload {
  status: "ok";
  authMode: "msal" | "personal";
  credential: PersonalHealth | { applicable: false };
  accounts: {
    path: string;
    count: number;
    activeId: string | null;
    repository: "registry" | "builtin";
    problem?: { code: string; detail: string; hint: string };
  };
}

export async function buildHealthzPayload(deps: { probeCdp?: (url: string) => Promise<boolean> } = {}): Promise<HealthzPayload> {
  const accounts = safeRegistryStatus();
  if (!isPersonalMode()) {
    return { status: "ok", authMode: "msal", credential: { applicable: false }, accounts };
  }
  const health = await personalHealth(deps.probeCdp ? { probeCdp: deps.probeCdp } : {});
  return { status: "ok", authMode: "personal", credential: health, accounts };
}

function safeRegistryStatus(): HealthzPayload["accounts"] {
  try {
    const st = registryStatus();
    return {
      path: st.path,
      count: st.count,
      activeId: st.activeId,
      repository: st.repository,
      ...(st.problem ? { problem: { code: st.problem.code, detail: st.problem.detail, hint: st.problem.hint } } : {}),
    };
  } catch (err) {
    return {
      path: "",
      count: 0,
      activeId: null,
      repository: "builtin",
      problem: { code: "unreadable", detail: err instanceof Error ? err.message : String(err), hint: "check accounts.json" },
    };
  }
}
