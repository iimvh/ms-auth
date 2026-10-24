
export type PriorityAccessWindow = "day" | "week";

export interface PriorityAccessExhaustion {
  window: PriorityAccessWindow;
  model?: string;
  resetsAt: Date;
  message: string;
}

const PRIORITY_ACCESS_RE =
  /(?:you.?ve|you have)\s+used\s+(?:up\s+)?(?:all\s+of\s+)?your\s+(?:available\s+)?priority\s+access(?:\s+to\s+the\s+(?<model>[\w.\- ]{1,32}?)\s+model)?\s+for\s+(?:the\s+)?(?<window>today|day|week|this\s+week)\b/i;

function nextUtcMidnight(from: Date): Date {
  return new Date(Date.UTC(from.getUTCFullYear(), from.getUTCMonth(), from.getUTCDate() + 1));
}

function nextUtcMonday(from: Date): Date {
  const daysUntilMonday = ((8 - from.getUTCDay()) % 7) || 7;
  return new Date(
    Date.UTC(from.getUTCFullYear(), from.getUTCMonth(), from.getUTCDate() + daysUntilMonday),
  );
}

export function parsePriorityAccessExhaustion(
  text: string | null | undefined,
  now: Date = new Date(),
): PriorityAccessExhaustion | null {
  if (!text) return null;
  const m = text.match(PRIORITY_ACCESS_RE);
  if (!m) return null;
  const raw = (m.groups?.window ?? "").toLowerCase();
  const window: PriorityAccessWindow = raw.includes("week") ? "week" : "day";
  return {
    window,
    model: m.groups?.model?.trim() || undefined,
    resetsAt: window === "week" ? nextUtcMonday(now) : nextUtcMidnight(now),
    message: text.trim(),
  };
}

const GATE_PHRASES = [
  "you've used your available priority access",
  "you've used up all of your priority access",
  "you have used your available priority access",
  "you have used up all of your priority access",
];
const GATE_MAX = Math.max(...GATE_PHRASES.map((p) => p.length));

export function couldBePriorityAccessPrefix(text: string): boolean {
  const t = text
    .replace(/[\u2018\u2019]/g, "'")
    .trimStart()
    .toLowerCase()
    .slice(0, GATE_MAX);
  if (!t) return true;
  return GATE_PHRASES.some((p) => p.startsWith(t) || t.startsWith(p));
}

export function secondsUntilReset(
  exhaustion: PriorityAccessExhaustion,
  now: Date = new Date(),
): number {
  return Math.max(1, Math.ceil((exhaustion.resetsAt.getTime() - now.getTime()) / 1000));
}

export type MeteredBudget = "opus" | "sonnet-5.5";

export const METERED_BUDGETS: Readonly<Record<MeteredBudget, { label: string; usageKey: string; daily: string; weekly: string }>> = {
  opus: { label: "Opus", usageKey: "opus", daily: "ClaudeOpusQueryDaily", weekly: "ClaudeOpusQuery75" },
  "sonnet-5.5": { label: "Sonnet 5.5", usageKey: "sonnet55", daily: "ClaudeSonnet55QueryDaily", weekly: "ClaudeSonnet55QueryWeekly" },
};

export function parseMetering(raw: unknown): Record<string, number> | null {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return null;
  const out: Record<string, number> = {};
  for (const [k, v] of Object.entries(raw as Record<string, unknown>)) {
    const n = (v as { remainingAllowance?: unknown } | null)?.remainingAllowance;
    if (typeof n === "number" && Number.isFinite(n)) out[k] = n;
  }
  return Object.keys(out).length ? out : null;
}

export function meteredAllowance(
  metering: Record<string, number> | null | undefined,
  budget: MeteredBudget,
): { daily?: number; weekly?: number } | null {
  if (!metering) return null;
  const daily = metering[METERED_BUDGETS[budget].daily];
  const weekly = metering[METERED_BUDGETS[budget].weekly];
  if (daily === undefined && weekly === undefined) return null;
  return { ...(daily !== undefined ? { daily } : {}), ...(weekly !== undefined ? { weekly } : {}) };
}

export function priorityAccessExhaustionOf(
  result: { value: string; message?: string } | null | undefined,
  text: string | null | undefined,
  now: Date = new Date(),
): PriorityAccessExhaustion | null {
  const fromText = parsePriorityAccessExhaustion(result?.message, now) ?? parsePriorityAccessExhaustion(text, now);
  if (fromText) return fromText;
  if (result?.value !== "OutOfCredits") return null;
  const message = (result.message ?? text ?? "").trim() || "OutOfCredits";
  const window: PriorityAccessWindow = /\bweek|monday/i.test(message) ? "week" : "day";
  return { window, resetsAt: window === "week" ? nextUtcMonday(now) : nextUtcMidnight(now), message };
}

const exhaustions = new Map<string, PriorityAccessExhaustion>();

export function notePriorityAccessExhausted(key: string, e: PriorityAccessExhaustion, now: Date = new Date()): void {
  const known = exhaustions.get(key);
  if (!known || e.resetsAt > known.resetsAt || known.resetsAt <= now) exhaustions.set(key, e);
}

export function activePriorityAccessExhaustion(key: string, now: Date = new Date()): PriorityAccessExhaustion | null {
  const known = exhaustions.get(key);
  if (known && known.resetsAt <= now) exhaustions.delete(key);
  return exhaustions.get(key) ?? null;
}

export function resetPriorityAccessState(): void {
  exhaustions.clear();
}
