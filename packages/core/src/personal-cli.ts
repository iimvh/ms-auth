import "./personal-wiring.js";
import { statSync, rmSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { pathToFileURL } from "node:url";
import { capturePersonalSession } from "./adapters/personal/personal-capture.js";
import type { PersonalCapture, PersonalRecord } from "./adapters/personal/personal-session.js";
import { PersonalStore } from "./adapters/personal/personal-session.js";
import { AccountError, accountsDir, configDir, resolveAccount } from "./account-registry.js";
import { probePersonalCapture } from "./session.js";
import { createLogger } from "./log.js";

const log = createLogger("personal-cli");

const USAGE = `m365-personal — manage the personal-account M365 session

usage: m365-personal <command> [--account <id>] [--json]

  login    capture the web app's session in a visible browser and store it
  status   print the stored session's expiry, scenario and path
  logout   delete the stored session and the browser profile

  --account <id>  act on that account instead of the active one (the ids are
                  the proxy's /accounts ones; the default account keeps the
                  pre-registry layout)
`;

export interface CliDeps {
  store: PersonalStore;
  /** The account acted on; null = the builtin layout (one account, no registry). */
  accountId: string | null;
  /** Where that account's Chrome profile lives (logout deletes it). */
  profileDir: string;
  capture: (headed: boolean) => Promise<PersonalCapture>;
  probe: (capture: PersonalCapture) => Promise<PersonalRecord["probe"]>;
  out: (line: string) => void;
  now: () => number;
}

/** Candidate ladder for the login-time capability probe (spec §10.1). */
function probeModels(): string[] {
  return (process.env.M365_PERSONAL_PROBE_MODELS ?? "claude-sonnet")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

function defaultDeps(accountId: string | null): CliDeps {
  const resolved = resolveAccount({ id: accountId });
  return {
    accountId: resolved.id,
    profileDir: resolved.profileDir,
    store: new PersonalStore(resolved.sessionFile),
    capture: (headed) =>
      capturePersonalSession({ headed, profileDir: resolved.profileDir, onEvent: (line) => console.log(line) }),
    probe: (capture) =>
      process.env.M365_PERSONAL_PROBE === "0"
        ? Promise.resolve(null)
        : probePersonalCapture(capture, probeModels()),
    out: (line) => console.log(line),
    now: () => Date.now(),
  };
}

/** Everything except the bearer token — `accessToken` never reaches stdout. */
function publicRecord(record: PersonalRecord) {
  const { accessToken, ...rest } = record;
  void accessToken;
  return rest;
}

function expiryLines(record: PersonalRecord, now: number): string[] {
  const expires = new Date(record.expiresAt);
  const minutes = Math.round((record.expiresAt - now) / 60_000);
  return [
    `expires: ${expires.toISOString()} (${minutes >= 0 ? "in " : "OVERDUE by "}${Math.abs(minutes)} min)`,
    `scenario: ${record.scenario}`,
    `chathubPath: ${record.chathubPath}`,
  ];
}

function profileExists(profileDir: string): boolean {
  try {
    return statSync(profileDir).isDirectory();
  } catch {
    return false;
  }
}

function usageError(why: string, out?: (line: string) => void): number {
  console.error(why);
  console.error(USAGE);
  // Also through `out`, so a caller that shows the CLI's own output (a wrapper,
  // a test) sees the reason rather than an exit code alone.
  out?.(why);
  return 2;
}

/**
 * Pull `--account <id>` out of the flags. Anything else is unknown. A missing
 * value, an unusable id, or an id the registry does not hold is a usage error
 * (exit 2) raised BEFORE any capture or write — naming the account must never
 * be able to fall back to the active one.
 */
function parseFlags(flags: string[], known: string[]): { ok: true; accountId: string | null } | { ok: false; why: string } {
  let accountId: string | null = null;
  for (let i = 0; i < flags.length; i++) {
    const f = flags[i]!;
    if (f === "--account") {
      const v = flags[i + 1];
      if (!v || v.startsWith("--")) return { ok: false, why: "--account needs an account id" };
      accountId = v;
      i++;
      continue;
    }
    if (f.startsWith("--account=")) {
      accountId = f.slice("--account=".length);
      if (!accountId) return { ok: false, why: "--account needs an account id" };
      continue;
    }
    if (!known.includes(f)) return { ok: false, why: `unknown flag ${f}` };
  }
  if (accountId !== null) {
    try {
      resolveAccount({ id: accountId });
    } catch (err) {
      return { ok: false, why: err instanceof AccountError ? err.message : String(err) };
    }
  }
  return { ok: true, accountId };
}

/** Run the `m365-personal` subcommand. Returns the process exit code. */
export async function runPersonalCli(argv: string[], deps?: Partial<CliDeps>): Promise<number> {
  const [command, ...flags] = argv;
  const json = flags.includes("--json");
  // --account is resolved BEFORE anything is captured or written: naming an
  // account must never silently fall back to the active one (§7.9).
  // Only the FIRST --json is dropped, so `--account --json` cannot leave the
  // flag's own name to be read as an account id.
  const jsonAt = flags.indexOf("--json");
  const flagArgs = jsonAt === -1 ? flags : [...flags.slice(0, jsonAt), ...flags.slice(jsonAt + 1)];
  const parsed = parseFlags(flagArgs, []);
  if (!parsed.ok) return usageError(parsed.why, deps?.out);
  let d: CliDeps;
  try {
    d = { ...defaultDeps(parsed.accountId), ...deps };
  } catch (err) {
    return usageError(err instanceof AccountError ? err.message : String(err), deps?.out);
  }
  const who = d.accountId ? `account "${d.accountId}"` : "the personal account";
  const loginCommand = `pnpm m365-personal login${d.accountId ? ` --account ${d.accountId}` : ""}`;

  switch (command) {
    case "login": {
      d.out(`capturing the personal session for ${who} in a visible browser...`);
      try {
        const capture = await d.capture(true);
        const probe = await d.probe(capture);
        const record: PersonalRecord = {
          ...capture,
          emailHint: null,
          probe: probe ?? null,
          mintCount: 1,
          lastMintAt: d.now(),
        };
        d.store.write(record);
        log.debug(`login captured account=${d.accountId ?? "builtin"} scenario=${record.scenario} mintCount=${record.mintCount}`);
        d.out(`account: ${d.accountId ?? "(builtin layout — the only account)"}`);
        for (const line of expiryLines(record, d.now())) d.out(line);
        if (probe && probe.ok) {
          d.out(`probed model: ${probe.model}`);
          if (json) d.out(JSON.stringify(publicRecord(record)));
          return 0;
        }
        if (json) d.out(JSON.stringify(publicRecord(record)));
        if (probe && !probe.ok) {
          // Stop condition (spec §10.1): never report success for a capture the
          // proxy cannot use. The record is still written — a login that gets an
          // answer later re-probes and replaces it.
          d.out("the structured scenario does not serve this account: no candidate model answered the capability probe");
          return 3;
        }
        d.out("capability probe: skipped");
        return 0;
      } catch (err) {
        console.error(`login failed: ${err instanceof Error ? err.message : String(err)}`);
        console.error(`run: ${loginCommand}`);
        return 1;
      }
    }

    case "status": {
      const state = d.store.inspect();
      if (!state.ok) {
        // "Nothing stored" and "what is stored is broken" need different repairs;
        // reporting both as the former sent the operator to `login` blind.
        if (state.problem === "absent") {
          d.out(`no personal session stored for ${who}.`);
          d.out(`run: ${loginCommand}`);
          return 1;
        }
        d.out(`the personal session record for ${who} is unusable (${state.problem}): ${state.detail}`);
        d.out(`delete it (or run: pnpm m365-personal logout${d.accountId ? ` --account ${d.accountId}` : ""}) and run: ${loginCommand}`);
        if (json) d.out(JSON.stringify({ account: d.accountId, problem: state.problem, detail: state.detail, profileExists: profileExists(d.profileDir) }));
        return 1;
      }
      const record = state.record;
      if (json) {
        d.out(JSON.stringify({ ...publicRecord(record), account: d.accountId, profileExists: profileExists(d.profileDir) }));
        return 0;
      }
      d.out(`account: ${d.accountId ?? "(builtin layout — the only account)"}`);
      for (const line of expiryLines(record, d.now())) d.out(line);
      d.out(`account hint: ${record.emailHint ?? "(none recorded)"}`);
      d.out(`browser profile present: ${profileExists(d.profileDir) ? "yes" : "no"}`);
      return 0;
    }
    case "logout": {
      // Deleting the profile is `rm -rf`, and with a registry the profile dir
      // comes from the file: a hand-edited entry pointing at the data dir
      // itself would take secrets.json and every other account with it. The
      // same guard as removeAccount, and the same refusal.
      const profile = resolve(d.profileDir);
      const config = resolve(configDir());
      if (profile === config || profile === resolve(accountsDir()) || profile === dirname(config)) {
        throw new AccountError(
          "registry_invalid",
          `refusing to delete ${profile}: it is the data dir itself, not one account's browser profile`,
          "fix that account's profileDir in accounts.json (it should be <data dir>/accounts/<id>/profile)",
        );
      }
      d.store.delete();
      try {
        rmSync(d.profileDir, { recursive: true, force: true });
      } catch (err) {
        log.debug("deleting the personal browser profile failed:", err);
      }
      d.out(`deleted the stored session and the browser profile for ${who}.`);
      return 0;
    }

    default:
      return usageError(command ? `unknown command ${command}` : "usage", deps?.out);
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.exitCode = await runPersonalCli(process.argv.slice(2));
}
