import type { PersonalRecord } from "./personal-session.js";
import { PersonalStore, personalSessionFile } from "./personal-session.js";

export type AuthMode = "msal" | "personal";

/** Which auth path this process runs. Defaults to today's enterprise MSAL;
 *  an unknown value fails here, so everything that reads it at boot can name
 *  the allowed values (Global Constraints). */
export function resolveAuthMode(): AuthMode {
  const raw = process.env.M365_AUTH ?? "msal";
  const mode = raw.toLowerCase();
  if (mode === "msal" || mode === "personal") return mode;
  throw new Error(`M365_AUTH must be "msal" or "personal", got "${raw}"`);
}

export function isPersonalMode(): boolean {
  return resolveAuthMode() === "personal";
}

/** Personal mode turns off everything an MSA token cannot serve (spec §10.4). */
export function personalModeDefaults(): { useAgent: false; codeInterpreter: boolean; imageGeneration: false } {
  return {
    useAgent: false,
    codeInterpreter: process.env.M365_PERSONAL_CODE_INTERPRETER === "1",
    imageGeneration: false,
  };
}

/** Code-interpreter optionsSets are what pull agent-less models into M365's
 *  own sandbox (§10.4): send them on the agent-less path — never on the agent
 *  path — except in personal mode, where the sandbox stays off unless the
 *  operator passes M365_PERSONAL_CODE_INTERPRETER=1. */
export function shouldSendCodeInterpreter(agentId: string | null | undefined, personal?: boolean): boolean {
  const isPersonal = personal ?? isPersonalMode();
  if (agentId != null) return false;
  if (isPersonal) return process.env.M365_PERSONAL_CODE_INTERPRETER === "1";
  return true;
}

/** The default upstream model in personal mode: `M365_PERSONAL_MODEL` wins,
 *  then the model the login-time capability probe got an answer from — the
 *  explicit argument (the CLI calls it with the record it just wrote) or, in
 *  the proxy path, the probe stored in the session record (spec §10.1). */
export function personalDefaultModel(
  probe?: { model: string; ok: boolean; at: number } | null,
): string | null {
  if (process.env.M365_PERSONAL_MODEL) return process.env.M365_PERSONAL_MODEL;
  const stored: { model: string; ok: boolean; at: number } | null = probe
    ?? (new PersonalStore(personalSessionFile()).read() as PersonalRecord | null)?.probe
    ?? null;
  if (stored && stored.ok) return stored.model;
  return null;
}
