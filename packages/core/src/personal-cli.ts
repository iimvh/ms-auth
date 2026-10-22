import { statSync, rmSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { capturePersonalSession } from "./personal-capture.js";
import type { PersonalCapture, PersonalRecord } from "./personal-session.js";
import { PersonalStore, personalProfileDir, personalSessionFile } from "./personal-session.js";
import { probePersonalCapture } from "./session.js";
import { createLogger } from "./log.js";

const log = createLogger("personal-cli");

const LOGIN_COMMAND = "pnpm m365-personal login";

const USAGE = `m365-personal — manage the personal-account M365 session

usage: m365-personal <command> [--json]

  login    capture the web app's session in a visible browser and store it
  status   print the stored session's expiry, scenario and path
  logout   delete the stored session and the browser profile
`;

export interface CliDeps {
  store: PersonalStore;
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

function defaultDeps(): CliDeps {
  return {
    store: new PersonalStore(personalSessionFile()),
    capture: (headed) =>
      capturePersonalSession({ headed, profileDir: personalProfileDir(), onEvent: (line) => console.log(line) }),
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

function profileExists(): boolean {
  try {
    return statSync(personalProfileDir()).isDirectory();
  } catch {
    return false;
  }
}

function usageError(why: string): number {
  console.error(why);
  console.error(USAGE);
  return 2;
}

/** Run the `m365-personal` subcommand. Returns the process exit code. */
export async function runPersonalCli(argv: string[], deps?: Partial<CliDeps>): Promise<number> {
  const d: CliDeps = { ...defaultDeps(), ...deps };
  const [command, ...flags] = argv;
  const json = flags.includes("--json");
  const unknownFlags = flags.filter((f) => f !== "--json");

  switch (command) {
    case "login": {
      if (unknownFlags.length > 0) return usageError(`unknown flag ${unknownFlags[0]}`);
      d.out("capturing the personal session in a visible browser…");
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
        log.debug(`login captured scenario=${record.scenario} mintCount=${record.mintCount}`);
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
        console.error(`run: ${LOGIN_COMMAND}`);
        return 1;
      }
    }

    case "status": {
      if (unknownFlags.length > 0) return usageError(`unknown flag ${unknownFlags[0]}`);
      const record = d.store.read();
      if (!record) {
        d.out("no personal session stored.");
        d.out(`run: ${LOGIN_COMMAND}`);
        return 1;
      }
      if (json) {
        d.out(JSON.stringify({ ...publicRecord(record), profileExists: profileExists() }));
        return 0;
      }
      for (const line of expiryLines(record, d.now())) d.out(line);
      d.out(`account hint: ${record.emailHint ?? "(none recorded)"}`);
      d.out(`browser profile present: ${profileExists() ? "yes" : "no"}`);
      return 0;
    }

    case "logout": {
      d.store.delete();
      try {
        rmSync(personalProfileDir(), { recursive: true, force: true });
      } catch (err) {
        log.debug("deleting the personal browser profile failed:", err);
      }
      d.out("deleted the stored session and the browser profile.");
      return 0;
    }

    default:
      return usageError(command ? `unknown command ${command}` : "usage");
  }
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  process.exitCode = await runPersonalCli(process.argv.slice(2));
}
