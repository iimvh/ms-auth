import { z } from "zod/v4";

export const PERSONAL_MINT_RATE_LIMIT_ENV = {
  minIntervalMs: "M365_PERSONAL_MINT_MIN_INTERVAL_MS",
  maxPerHour: "M365_PERSONAL_MINT_MAX_PER_HOUR",
} as const;

export const PERSONAL_MINT_RATE_LIMITS = {
  minIntervalMs: 600_000,
  maxPerHour: 3,
} as const;

export type PersonalMintRateLimit = keyof typeof PERSONAL_MINT_RATE_LIMITS;

export const MINT_HOUR_MS = 3_600_000;

export const PersonalMintRateLimitOverrideSchema = z.object({
  minIntervalMs: z.number().nonnegative().optional(),
  maxPerHour: z.number().nonnegative().optional(),
});

export type PersonalMintRateLimitOverride = z.infer<typeof PersonalMintRateLimitOverrideSchema>;

function envNumber(env: Record<string, string | undefined>, limit: PersonalMintRateLimit): number | undefined {
  const raw = env[PERSONAL_MINT_RATE_LIMIT_ENV[limit]];
  if (raw === undefined) return undefined;
  const value = Number(raw);
  return Number.isFinite(value) ? value : undefined;
}

export function personalMintRateLimits(
  env: Record<string, string | undefined> = process.env,
  override: PersonalMintRateLimitOverride = {},
): Record<PersonalMintRateLimit, number> {
  const parsed = PersonalMintRateLimitOverrideSchema.parse(override);
  return {
    minIntervalMs: parsed.minIntervalMs ?? envNumber(env, "minIntervalMs") ?? PERSONAL_MINT_RATE_LIMITS.minIntervalMs,
    maxPerHour: parsed.maxPerHour ?? envNumber(env, "maxPerHour") ?? PERSONAL_MINT_RATE_LIMITS.maxPerHour,
  };
}
