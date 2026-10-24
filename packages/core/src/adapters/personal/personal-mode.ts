import type { PersonalRecord } from "./personal-session.js";
import { PersonalStore } from "./personal-session.js";
import { resolveAccount } from "../../account-registry.js";

export type AuthMode = "msal" | "personal";

export function resolveAuthMode(): AuthMode {
  const raw = process.env.M365_AUTH ?? "msal";
  const mode = raw.toLowerCase();
  if (mode === "msal" || mode === "personal") return mode;
  throw new Error(`M365_AUTH must be "msal" or "personal", got "${raw}"`);
}

export function isPersonalMode(): boolean {
  return resolveAuthMode() === "personal";
}

export function personalModeDefaults(): { useAgent: false; codeInterpreter: boolean; imageGeneration: false } {
  return {
    useAgent: false,
    codeInterpreter: process.env.M365_PERSONAL_CODE_INTERPRETER === "1",
    imageGeneration: false,
  };
}

export function shouldSendCodeInterpreter(agentId: string | null | undefined, personal?: boolean): boolean {
  const isPersonal = personal ?? isPersonalMode();
  if (agentId != null) return false;
  if (isPersonal) return process.env.M365_PERSONAL_CODE_INTERPRETER === "1";
  return true;
}

export function personalDefaultModel(
  probe?: { model: string; ok: boolean; at: number } | null,
): string | null {
  if (process.env.M365_PERSONAL_MODEL) return process.env.M365_PERSONAL_MODEL;
  const stored: { model: string; ok: boolean; at: number } | null = probe
    ?? (new PersonalStore(resolveAccount().sessionFile).read() as PersonalRecord | null)?.probe
    ?? null;
  if (stored && stored.ok) return stored.model;
  return null;
}
