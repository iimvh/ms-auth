import { createLogger } from "../io/log.js";

const log = createLogger("backoff");

export interface BackoffOptions {
  now?: () => number;
  sleep?: (ms: number) => Promise<void>;
  rng?: () => number;
  windowMs?: number;
  threshold?: number;
  baseCooldownMs?: number;
  maxCooldownMs?: number;
  jitterMinMs?: number;
  jitterMaxMs?: number;
  onTrigger?: (info: { distinctConversations: number; cooldownMs: number; level: number }) => void;
}

export interface BackoffController {
  note: (empty: boolean, conversationId: string) => void;
  waitForSlot: () => Promise<number>;
  isBackingOff: () => boolean;
}

export function createBackoffController(opts: BackoffOptions): BackoffController {
  const now = opts.now ?? (() => Date.now());
  const sleep = opts.sleep ?? ((ms: number) => new Promise<void>((r) => setTimeout(r, ms)));
  const rng = opts.rng ?? Math.random;
  const windowMs = opts.windowMs ?? 120_000;
  const threshold = opts.threshold ?? 3;
  const baseCooldownMs = opts.baseCooldownMs ?? 90_000;
  const maxCooldownMs = opts.maxCooldownMs ?? 600_000;
  const jitterMinMs = opts.jitterMinMs ?? 10_000;
  const jitterMaxMs = opts.jitterMaxMs ?? 25_000;

  let empties: Array<{ t: number; conv: string }> = [];
  let backoffUntil = -Infinity;
  let level = 0;

  return {
    note(empty, conversationId) {
      const t = now();
      if (!empty) {
        empties = [];
        backoffUntil = -Infinity;
        level = 0;
        return;
      }

      empties.push({ t, conv: conversationId || `anon-${t}` });
      empties = empties.filter((e) => t - e.t < windowMs);

      const distinct = new Set(empties.map((e) => e.conv)).size;
      if (distinct < threshold) return;
      if (t < backoffUntil) return;

      level += 1;
      const cooldownMs = Math.min(baseCooldownMs * 2 ** (level - 1), maxCooldownMs);
      backoffUntil = t + cooldownMs;
      empties = [];
      opts.onTrigger?.({ distinctConversations: distinct, cooldownMs, level });
    },

    async waitForSlot() {
      const remaining = backoffUntil - now();
      if (remaining <= 0) return 0;
      const span = Math.max(0, jitterMaxMs - jitterMinMs);
      const delay = Math.min(remaining, jitterMinMs + Math.floor(rng() * (span + 1)));
      if (delay > 0) await sleep(delay);
      return delay;
    },

    isBackingOff() {
      return now() < backoffUntil;
    },
  };
}

const disabled = () =>
  !!(process.env.M365_NO_BACKOFF ?? process.env.M365_NO_AUTO_REAUTH);

function controllerOptions(account: string | null): BackoffOptions {
  return {
    windowMs: Number(process.env.M365_BACKOFF_WINDOW_MS ?? 120_000),
    threshold: Number(process.env.M365_BACKOFF_THRESHOLD ?? process.env.M365_REAUTH_EMPTY_THRESHOLD ?? 3),
    baseCooldownMs: Number(process.env.M365_BACKOFF_BASE_MS ?? 90_000),
    maxCooldownMs: Number(process.env.M365_BACKOFF_MAX_MS ?? 600_000),
    onTrigger: ({ distinctConversations, cooldownMs, level }) => {
      const whose = account ? `account "${account}"` : "the account";
      log.info(
        `Degradation backoff (level ${level}): ${distinctConversations} empty responses across distinct ` +
        `conversations on ${whose} — pacing new turns for ~${Math.round(cooldownMs / 1000)}s to let it self-heal ` +
        `(H-R1: a re-login would NOT clear this and would raise our detection profile). Disable with M365_NO_BACKOFF=1.`,
      );
    },
  };
}

const controllers = new Map<string, BackoffController>();

function controllerFor(accountId?: string | null): BackoffController {
  const key = accountId ?? "builtin";
  let c = controllers.get(key);
  if (!c) {
    c = createBackoffController(controllerOptions(accountId ?? null));
    controllers.set(key, c);
  }
  return c;
}

export function noteRequestOutcome(empty: boolean, conversationId: string, accountId?: string | null): void {
  if (disabled()) return;
  controllerFor(accountId).note(empty, conversationId);
}

export async function awaitDegradationBackoff(accountId?: string | null): Promise<void> {
  if (disabled()) return;
  const slept = await controllerFor(accountId).waitForSlot();
  if (slept > 0) log.info(`Backoff: paced this turn by ${Math.round(slept / 1000)}s (${accountId ? `account "${accountId}"` : "account"} degraded)`);
}

export function isDegradationBackoff(accountId?: string | null): boolean {
  return !disabled() && controllerFor(accountId).isBackingOff();
}

export function forgetAccountBackoff(accountId: string): void {
  controllers.delete(accountId);
}

export function resetBackoffState(): void {
  controllers.clear();
}
