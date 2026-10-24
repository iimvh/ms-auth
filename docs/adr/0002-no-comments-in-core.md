# 0002 — No comments in `packages/core` and `packages/proxy-lib`

**Status:** accepted
**See also:** `AGENTS.md` §"No Comments Policy"

## Context

The repository carries about 2,400 comment lines over 10,700 code lines in
`packages/core` and `packages/proxy-lib`. The density is not uniform, and the
worst cases are the files that change most:

| File | Total | Comment | Code |
|---|---|---|---|
| `core/src/fenced.ts` | 1475 | 464 | 858 |
| `core/src/copilot.ts` | 574 | 345 | 187 |
| `core/src/session.ts` | 1230 | 271 | 892 |
| `core/src/tools.ts` | 593 | 194 | 352 |
| `core/src/account-registry.ts` | 913 | 159 | 682 |
| `core/src/turn-budget.ts` | 491 | 141 | 317 |

`copilot.ts` has more comment than code. Most of it restates what a table or a
branch already says, narrates history that `docs/hypotheses.md` records with
measurements, or explains a shape that a type would state.

`AGENTS.md` bans the vast majority of comments and permits one rare case: a real
external constraint that cannot be expressed as a type, test, assertion or lint
rule. Applying that policy to 2,400 lines without a written record of which
comments were kept would make the result unverifiable. A deleted comment does not
fail a test and does not read as an error in a diff; it reads as an empty line.

## Decision

Delete comments in two phases.

Phase 1 removes every comment not listed in the allow-list below. No code line
changes, no assertion changes, and `pnpm test:unit` must stay green before and
after. This happens before any file moves, so a red test proves a comment was
load-bearing.

Phase 2 moves files according to ADR-0001. Comments are not reintroduced during
or after the move.

## Allow-list

These are kept, or converted, as stated. Everything else in
`packages/core/src` and `packages/proxy-lib/src` is deleted.

| Location | Treatment | Reason |
|---|---|---|
| `core/src/account-registry.ts:45–60` | keep | Defines the id grammar and the pre-accounts layout; neither is derivable from code |
| `core/src/account-registry.ts:136–173` | keep, moves to `testing.ts` | States what the mock registry deliberately does not emulate |
| `core/src/account-registry.ts:363–411` | keep | Describes the layout that existed before accounts existed |
| `core/src/account-registry.ts:796–812` | delete | Precedence is already asserted in `account-registry.test.ts:389–412` |
| `core/src/account-registry.ts:554` | delete | `PersonalRecordState` states it |
| `core/src/session.ts` header and `dumpFrame` | keep 3 lines | `M365_DUMP_FRAMES` writes to `FRAME_DIR`; the variable's effect is environmental |
| `core/src/session.ts` frame comments | delete after #2 and #3 land | Convert to `frames.ts` schemas and a replay test first |
| `core/src/fenced.ts:1–60` | delete | The measurement is in `README.md` and `docs/hypotheses.md` §9, and `fenced.test.ts` asserts the behaviour |
| `core/src/fenced.ts` variant docs | delete | The variant name is the documentation; see ADR-0001 |
| `core/src/copilot.ts:1–150` | keep one block, delete the rest | One block states what the tone table is; the table is self-describing |
| `core/src/copilot.ts:300–380` | delete | Which set Microsoft defines versus which we choose belongs in a test |
| `core/src/turn-budget.ts:1–40` | keep one cross-reference | The throttle formula is not derivable from code; it is written up in `docs/hypotheses.md` |
| `core/src/auth.ts:32`, `auth.ts:270–290` | keep | The artifact endpoint is undocumented by Microsoft |
| `core/src/auth.ts:550–555` | delete after an assertion | The `pollForToken` hang is assertable |
| `core/src/auth.ts:630–660` | delete after a schema | The secrets file has a shape; express it in zod |
| `core/src/personal-capture.ts:92` | keep one cross-reference | The TTL is an estimate, not a measurement; the numbers are in `docs/personal-mint-design.md` |
| `core/src/personal-capture.ts:340–420` | keep | A Playwright-launched browser cannot pass the Cloudflare gate; the constraint is Microsoft's |
| `core/src/personal-auth.ts:11–50` | delete | `PersonalAuthErrorCode` already states the codes |
| `core/src/personal-auth.ts:150–360` | delete | Rate limits are named constants with env defaults |
| `core/src/personal-session.ts:63–78` | delete | `PersonalRecordProblem` and `PersonalRecordState` are union types |
| `core/src/personal-harvest.ts:1–50` | keep | The page opens its socket only when a message is sent; measured, undocumented |
| `core/src/personal-mode.ts` | delete all | Logic selecting a mode |
| `core/src/personal-cli.ts` | delete all | Logic |
| `core/src/personal-health.ts` | delete all | Logic |
| `proxy-lib/src/handler.ts:150–200` | delete | Session pool keying is asserted in `handler-account.test.ts` |
| `proxy-lib/src/handler.ts:440–470` | keep 2 lines | Metering is a property of the paid scenarios; a rule of the upstream API |
| `proxy-lib/src/handler.ts:540–580` | delete | Logic |
| `proxy-lib/src/handler.ts:660–700` | keep | Disengagement is permanent; observed, undocumented |
| `proxy-lib/src/healthz.ts:5–35` | keep 1 line | `/healthz` must not mint or spend a turn; a design constraint |
| `proxy-lib/src/accounts-api.ts:26` | delete | The 404 is in the code path and will be covered by the proxy route tests |
| `packages/proxy/server-accounts-ui.ts:4–7` | keep | Nitro's build makes `readFileSync` unreliable; a platform constraint |
| `proxy-lib/src/index.ts:100–115` | delete | Logic |
| `proxy-lib/src/index.ts:151` | delete | Describes the surface |

## Consequences

- About 2,340 comment lines are deleted; about 14 blocks (roughly 60 lines) remain.
- The six prose-only behaviours now have code (#16): precedence in
  `account-registry.test.ts`, record state pinned by `PersonalRecordProblemSchema`
  and `PersonalRecordState`, error codes by `PERSONAL_AUTH_ERROR_CODES`, rate
  limits by `personal-mint-limits.ts`, the secrets file shape by
  `SecretsFileSchema`, and the `pollForToken` hang by an assertion in
  `auth.test.ts`. Every "delete after X" row on those has become a plain delete.
- Frame-format comments in `session.ts` cannot be deleted until recorded frame
  fixtures exist and a replay test covers them (#2, #3). Until then they are
  converted, not removed, and that conversion belongs to the fixtures task.
- Reviewing the deletion is possible because the allow-list names every survivor.
  Reviewing it by reading the diff is not.

## Rejected alternatives

**Delete everything, including the allow-list.** Violates the single exception
`AGENTS.md` grants, and removing the note about Nitro's build leads the next
reader to add a `readFileSync` that 404s under a different launch directory.

**Delete during the file moves.** Makes a 25-file diff that both relocates and
loses 464 lines in one file, so neither change can be reviewed on its own.

**Defer to after the refactor.** Leaves `fenced.ts` at 1475 lines while it is
being moved, which is the worst time to have 464 lines of churn.
