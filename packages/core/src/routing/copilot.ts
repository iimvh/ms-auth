import { JwtClaims } from "../protocol/schemas.js";
import type { MeteredBudget } from "../budget/priority-access.js";

// Model name → tone mapping.
// The server VALIDATES tones (an unknown tone errors with "Failed to invoke
// 'Chat'"), so every entry here has been confirmed accepted against the live
// API. Claude tones self-identify as Anthropic models (docs/hypotheses.md
// H8.6) — a genuine non-Microsoft model at zero marginal cost.
const MODEL_TONES: Record<string, string> = {
  "m365-copilot": "magic",
  "auto": "magic",

  "quick": "Gpt_5_5_Chat",
  "think-deeper": "Gpt_5_5_Reasoning",

  "claude": "Claude_Sonnet",
  "claude-sonnet": "Claude_Sonnet",
  "claude-sonnet-4.5": "Claude_Sonnet",
  "claude-sonnet-4.6": "Claude_Sonnet",
  "claude-sonnet-5.5": "Claude_Sonnet",
  "claude-sonnet-5": "Claude_Sonnet_5",
  "claude-sonnet-think-deeper": "Claude_Sonnet_Reasoning",
  "claude-opus": "Claude_Opus",
  "claude-opus-5": "Claude_Opus",
  "claude-opus-5.5": "Claude_Opus",
  "claude-opus-4.8": "Claude_Opus",
  "claude-opus-4.5": "Claude_Opus",

  "gpt-5.5": "Gpt_5_5_Chat",
  "gpt-5.5-quick": "Gpt_5_5_Chat",
  "gpt-5.5-think-deeper": "Gpt_5_5_Reasoning",

  "gpt-5.6": "Gpt_5_6_Chat",
  "gpt-5.6-quick": "Gpt_5_6_Chat",
  "gpt-5.6-think-deeper": "Gpt_5_6_Reasoning",

  "gpt-6-think-deeper": "Gpt_6_Reasoning",

  "gpt-6-sol": "Gpt_6_Sol_Reasoning",

  "gpt-6.1-sol": "Gpt_61_Sol_Reasoning",

  "gpt-5.4": "Gpt_5_4_Reasoning",
  "gpt-5.4-think-deeper": "Gpt_5_4_Reasoning",
  "gpt-5.4-quick": "Gpt_5_4_Chat",

  "gpt-5.3": "Gpt_5_3_Chat",
  "gpt-5.3-quick": "Gpt_5_3_Chat",
  "gpt-5.3-think-deeper": "Gpt_5_3_Reasoning",

  "gpt-5.2": "Gpt_5_2_Chat",
  "gpt-5.2-quick": "Gpt_5_2_Chat",
  "gpt-5.2-think-deeper": "Gpt_5_2_Reasoning",
};

export function getToneForModel(model: string): string {
  const exact = MODEL_TONES[model];
  if (exact) return exact;
  if (/opus/i.test(model)) return "Claude_Opus";
  if (/^claude/i.test(model)) {
    return SONNET_5_PATTERN.test(model) && !SONNET_5_5_PATTERN.test(model) ? "Claude_Sonnet_5" : "Claude_Sonnet";
  }
  return MODEL_TONES["m365-copilot"];
}

const DEFAULT_SCENARIO = "OfficeWebIncludedCopilot";
const DEFAULT_LICENSE_TYPE = "Starter";
const PAID_SCENARIO = "OfficeWebPaidCopilot";
const PAID_LICENSE_TYPE = "Premium";

export const PAID_SCENARIO_TONES: ReadonlySet<string> = new Set([
  "Claude_Opus",
  "Claude_Sonnet_5",
  "Gpt_6_Reasoning",
]);

export const PAID_SCENARIO_MODELS: ReadonlySet<string> = new Set([
  "claude-sonnet-5.5",
]);
const SONNET_5_5_PATTERN = /sonnet[-_ ]?5[-_. ]5(?!\d)/i;
const SONNET_5_PATTERN = /sonnet[-_ ]?5(?!\d|\.\d)/i;

export function isSonnet55Model(model: string): boolean {
  return PAID_SCENARIO_MODELS.has(model) ||
    (getToneForModel(model) === "Claude_Sonnet" && SONNET_5_5_PATTERN.test(model));
}

export function isSonnet5Model(model: string): boolean {
  return getToneForModel(model) === "Claude_Sonnet_5";
}

export const INCLUDED_SCENARIO_MODELS: ReadonlySet<string> = new Set([
  "claude-opus-4.8",
  "claude-opus-4.5",
]);
const INCLUDED_OPUS_PATTERN = /opus[-_ ]?4[-_. ][58](?!\d)/i;

export function isIncludedOpusModel(model: string): boolean {
  return INCLUDED_SCENARIO_MODELS.has(model) ||
    (getToneForModel(model) === "Claude_Opus" && INCLUDED_OPUS_PATTERN.test(model));
}

export const AGENTLESS_TOOL_TONES: ReadonlySet<string> = new Set([
  "Gpt_6_Reasoning",
]);

export const AGENT_CLAUDE_TONES: ReadonlySet<string> = new Set([
  "Claude_Opus",
]);

export const PREMIUM_ONLY_AGENT_TONES: ReadonlySet<string> = new Set([
  "Gpt_6_Sol_Reasoning",
  "Gpt_61_Sol_Reasoning",
]);

const deadAgentRoutes = new Set<string>();
const aliveAgentRoutes = new Set<string>();

export function noteAgentRouteDead(tone: string): boolean {
  if (aliveAgentRoutes.has(tone)) return false;
  deadAgentRoutes.add(tone);
  return true;
}

export function noteAgentRouteAlive(tone: string): void {
  aliveAgentRoutes.add(tone);
}

export function isAgentRouteAlive(tone: string): boolean {
  return aliveAgentRoutes.has(tone);
}

export function resetAgentRoutes(): void {
  deadAgentRoutes.clear();
  aliveAgentRoutes.clear();
}

export function toneUsesToolAgent(tone: string): boolean {
  if (AGENTLESS_TOOL_TONES.has(tone) || deadAgentRoutes.has(tone)) return false;
  if (/^Claude_/i.test(tone)) return AGENT_CLAUDE_TONES.has(tone);
  return true;
}

export function toolRequestUsesAgent(tone: string): boolean {
  const force = process.env.M365_FORCE_AGENT;
  if (force === "1") return true;
  if (force === "0") return false;
  return toneUsesToolAgent(tone);
}

export function requestUsesAgent(model: string, hasTools: boolean): boolean {
  return hasTools && toolRequestUsesAgent(getToneForModel(model));
}

export function isPremiumOnlyModel(model: string): boolean {
  return isIncludedOpusModel(model);
}

export interface ScenarioRouting {
  scenario: string;
  licenseType: string;
}

function routing(paid: boolean): ScenarioRouting {
  return {
    scenario: process.env.M365_SCENARIO ?? (paid ? PAID_SCENARIO : DEFAULT_SCENARIO),
    licenseType: process.env.M365_LICENSE_TYPE ?? (paid ? PAID_LICENSE_TYPE : DEFAULT_LICENSE_TYPE),
  };
}

export function getScenarioForTone(tone: string): ScenarioRouting {
  return routing(PAID_SCENARIO_TONES.has(tone));
}

export function getScenarioForModel(model: string): ScenarioRouting {
  if (isIncludedOpusModel(model)) return routing(false);
  return routing(PAID_SCENARIO_TONES.has(getToneForModel(model)) || isSonnet55Model(model));
}

export function meteredBudgetOf(model: string): MeteredBudget | null {
  if (getScenarioForModel(model).scenario !== PAID_SCENARIO) return null;
  const tone = getToneForModel(model);
  if (tone === "Claude_Opus") return "opus";
  if (tone === "Claude_Sonnet") return "sonnet-5.5";
  return null;
}

export const PRIORITY_ACCESS_FALLBACK_ENV: Readonly<Record<MeteredBudget, string>> = {
  opus: "M365_OPUS_FALLBACK_MODEL",
  "sonnet-5.5": "M365_SONNET_FALLBACK_MODEL",
};

export function priorityAccessFallbackModel(budget: MeteredBudget): string | null {
  const v = process.env[PRIORITY_ACCESS_FALLBACK_ENV[budget]]?.trim();
  if (!v || meteredBudgetOf(v)) return null;
  return v;
}

export function getAvailableModels(): string[] {
  return Object.keys(MODEL_TONES);
}

export function decodeJwt(token: string) {
  const payload = token.split(".")[1];
  const padded = payload + "=".repeat((4 - (payload.length % 4)) % 4);
  const raw = JSON.parse(Buffer.from(padded, "base64").toString());
  return JwtClaims.parse(raw);
}

export interface CapturedImage {
  referenceUrls: string[];
  fileToken?: string;
  pollUrl?: string;
  size?: string;
  orientation?: string;
  status?: number;
}

export interface CopilotStream {
  [Symbol.asyncIterator](): AsyncIterator<string>;
  fullText: string;
  images: CapturedImage[];
  hasContent: boolean;
  throttle: { current: number; max: number } | null;
  contentOrigin?: string | null;
  messageType?: string | null;
  messageId?: string | null;
  scores?: Record<string, number> | null;
  turnCount?: number | null;
  turnState?: string | null;
  result?: { value: string; errorCode?: string; message?: string } | null;
  metering?: Record<string, number> | null;
  sawAction?: boolean;
}
