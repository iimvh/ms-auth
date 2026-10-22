# Personal-account auth (`M365_AUTH=personal`)

How the proxy runs against a **personal** Microsoft account (MSA / "consumer"
Copilot) instead of the enterprise M365 Copilot this proxy was built on.

## What this path is — and is not

It is **not** an OAuth client. It captures what the real web app
(`m365.cloud.microsoft/chat`) does and replays it:

- The sign-in sits behind **Cloudflare** ("verify you are human" / "Security
  check required"). A browser *launched* by Playwright can never pass it — the
  launch itself is the fingerprint — so `login` attaches to **your own Chrome**
  over CDP (`M365_PERSONAL_CDP_URL=http://127.0.0.1:9222`, Chrome started with
  `--remote-debugging-port=9222 --user-data-dir=<personal profile>`). Only then
  does the page send its `ChallengeProof` / `CloudflareTurnstile` frame.
- `pnpm m365-personal login` waits for the sign-in to finish (an empty profile
  opens the Microsoft sign-in page and the command waits — up to 5 min,
  `M365_PERSONAL_LOGIN_WAIT_MS`; if `secrets.json` exists its email/password/TOTP
  drive the form automatically), types one short message to force the Chathub
  socket open, and records that session: the socket URL's exact parameters
  (`scenario=OfficeWebFreeConsumerCopilot`, `variants`, `licenseType`, `agentHost`,
  `source`, the path `oid@tid`; the free consumer socket carries **no `product`**)
  plus the first chat invocation frame (`optionsSets`, `allowedMessageTypes`), and
  the access token's expiry.
- The consumer `access_token` is an **opaque JWE** — no readable claims — so the
  expiry is estimated (`M365_PERSONAL_TOKEN_TTL_MS`, default 1 h; the observed
  token lives 60 minutes).
- At runtime the proxy rebuilds the socket URL from those captured values, so
  a Microsoft param change is fixed by re-running one command — no table to edit.
- When the token is near expiry it is re-minted from the same browser profile,
  headless first, headed once as a fallback. On Windows the consumer SSO does
  **not** auto-continue in a cold headless browser, so a re-mint that needs the
  sign-in page fails and the turn names the login command; keep the Chrome from
  the login step running if you want silent re-mints.

Everything lives in the existing data dir (`~/.config/opencode-m365/`):

| File | Contents | Sensitivity |
|---|---|---|
| `personal-profile/` | browser profile: Microsoft session cookies | **high** — full account session |
| `personal-session.json` | token + captured params + mint bookkeeping | **high** — live bearer token |

## The commands

```sh
# with M365_PERSONAL_CDP_URL set, attaches to your Chrome (the only way past Cloudflare)
pnpm m365-personal login     # one-time capture + capability probe
pnpm m365-personal status    # expiry, scenario, chathubPath, account hint — never the token
pnpm m365-personal status --json   # same, machine-readable (no accessToken field)
pnpm m365-personal logout    # deletes BOTH the record and the browser profile
```

`login` ends with a **capability probe**: one "reply with the single word ok"
turn per candidate model (`M365_PERSONAL_PROBE_MODELS`, default `claude-sonnet`)
and records which model answered; personal mode's default upstream model is
that one. `M365_PERSONAL_PROBE=0` skips the probe. If no candidate answers,
`login` exits 3 and says so — a capture the proxy cannot use is never reported
as success (the record is still written; a later re-login replaces it).

## Environment

| Variable | Default | Meaning |
|---|---|---|
| `M365_AUTH` | `msal` | `personal` = the path above. Anything else fails at boot naming the allowed values. |
| `M365_PERSONAL_MODEL` | record's probe | Override the default upstream model in personal mode. |
| `M365_PERSONAL_PROBE_MODELS` | `claude-sonnet` | Comma-separated candidate ladder for the login probe. |
| `M365_PERSONAL_PROBE` | `1` | `0` skips the login-time model probe. |
| `M365_PERSONAL_MINT_MIN_INTERVAL_MS` | `600000` | Headless mints are at least this far apart. |
| `M365_PERSONAL_MINT_MAX_PER_HOUR` | `3` | Headless mint budget per hour (each mint is a fresh conversation — the thread-rate throttle's unit). |
| `M365_PERSONAL_MINT_WAIT_MS` | `10000` | How long a mint waits for the app's socket on page load before typing anything. |
| `M365_PERSONAL_LOGIN_WAIT_MS` | `300000` | How long `login` (headed) waits on the Microsoft sign-in page for you to finish signing in. |
| `M365_PERSONAL_CDP_URL` | unset | Attach to an already-running Chrome (`http://127.0.0.1:9222`) instead of launching one. Required to clear Cloudflare. |
| `M365_PERSONAL_AUTOMATED_LOGIN` | `1` | `0` = never drive the sign-in form with `secrets.json`; wait for a human. |
| `M365_PERSONAL_TOKEN_TTL_MS` | `3600000` | Estimated lifetime of the opaque JWE token (the observed one lives 60 min). |
| `M365_PERSONAL_VERIFY_WAIT_MS` | `300000` | How long a capture waits for a human to clear the verification/security-check dialog, then sends again. Nothing bypasses the challenge. |
| `M365_PERSONAL_CODE_INTERPRETER` | unset | `1` re-enables M365's Python sandbox optionsSets in personal mode. |
| `M365_PERSONAL_PROFILE` | `<data dir>/personal-profile` | Browser profile location. |
| `M365_PERSONAL_SESSION_FILE` | `<data dir>/personal-session.json` | Record location. |

## What personal mode turns off, and why

| Behaviour | Personal mode | Why |
|---|---|---|
| Tool agent / Copilot Studio | **off, forced** | An MSA token has no powerplatform scopes; the agent path cannot serve it (nothing to buy, this is capture-and-replay). Tool calling still happens agent-less via the fenced/shell routing that the Claude path already uses. |
| M365 Python sandbox (code interpreter) | **off** | Agent-less turns pull M365's own sandbox instead of our fenced tools; `M365_PERSONAL_CODE_INTERPRETER=1` re-enables it knowingly. |
| Image generation | **off** | The renderer needs an MSAL-only artifact token a personal session cannot mint. |
| Model tones | **`tone: "Magic"` only** | The consumer backend answers `InternalError` to any other tone on an otherwise identical turn (F108). The model name in your request is accepted but the answer comes from consumer Copilot. The turn also carries the consumer `clientInfo` (`ProductCategory: Chat`, `productEntryPoint: ChatPanel`, `deviceOS: Windows`). |
| MSAL priority-access wall | **not applied** | Opus/Sonnet priority budgets belong to the paid scenarios; a personal session has none. The upstream `CostQuota` (80 turns/day for a free account, reset 00:00 UTC) is the real ceiling — every probe and retry counts. |

Framing defaults are unchanged (framing is load-bearing; a consumer-specific
sweep is a measured follow-up, not phase 1).

## Failure messages

Every failure that needs a human quotes the command to run — this is the whole
recovery story:

- **`M365 personal session expired — run: pnpm m365-personal login`**
  — the record is stale and the profile is dead (a re-mint landed on a sign-in
  page, or the capture could not clear Microsoft's verification). Run `login`;
  the proxy recovers mid-day without a restart (it re-reads the record on every
  turn).
- **`could not re-mint the M365 personal session — run: pnpm m365-personal login`**
  — the capture failed while the record was stale (both the headless and the
  headed attempt) — transient network, the Cloudflare/security-check dialog, or
  a Microsoft param/flow change. Retry once; if it persists, repeat the CDP
  login (the consumer SSO does not auto-continue in a cold headless browser, so
  a mint without the operator's Chrome may only ever end here).
- **`refusing to mint ... mint_rate_limited`** — the mint rate limits hit. Not
  dead: usually means you ran a fresh-conversation sweep against the personal
  account. Wait for the interval/hour budget, or `login` by hand.
- `M365_AUTH=<bad>` fails at boot: `M365_AUTH must be "msal" or "personal", got "<value>"`.

A dead session surfaces through the OpenAI API as
`401 {"error":{..., "code":"m365_personal_session_dead"}}` **before any stream
byte** is committed — a `stream:true` client gets a real HTTP status, not a
mid-stream error.

## Security note

`personal-profile/` is a **full signed-in session for your whole Microsoft
account** — anyone with that directory can act as you. It is never inside the
repo, is `0o600`-protected where the platform allows it, and `logout` deletes
both it and the JSON record. On a shared machine, do not leave a personal
profile behind. Windows ACL hardening beyond default user-file ACLs is left
open; the reasonable-person rule is: shared host → `logout` right after use.

## Operator checklist (the live gates — run these, record the numbers)

Nothing below is claimed by the code; phase 1 ends honestly at "personal chat
works" until these have real numbers in `docs/hypotheses.md`.

```text
1. pnpm m365-personal login                   → note expiry, scenario, probed model
2. pnpm --filter @m365-copilot/proxy dev       → boots, logs "personal mode"
3. M365_AUTH=personal node scripts/proxy-verify.mjs --tools --multiturn --model=claude-sonnet
                                               → record: tool call? turns? dea score?
4. one real Claude Code session through the Python gateway: create a file, run it
                                               → record: turns, tool calls, Disengaged count
5. force a stale record (edit expires_at back 10 min) and run step 3 again
                                               → record: did it mint? how long? (headless SSO
                                                 does not auto-continue on a cold Windows
                                                 browser — expect the login command, keep the
                                                 Chrome from step 1 running to avoid it)
6. rm -rf the personal profile, run step 3     → record: exact error text, no hang
7. after 24 h: pnpm m365-personal status --json → record: mint count, token lifetime
```

Known numbers so far (F97–F109 in `docs/hypotheses.md`): a free consumer account
is `scenario=OfficeWebFreeConsumerCopilot`, `licenseType=Free`, **80 turns/day**
(`CostQuota`, reset 00:00 UTC); the token lives 60 minutes; a captured session
answers `/v1/chat/completions` and an Anthropic gateway in front of it end to
end. Do not bench on a free account without counting turns.

Also worth recording when they run: whether the consumer scenario produces a
**fenced tool call** at all (spec §10.1: if not, phase 1 ends at "personal chat
works, tool routing unproven" and the next step is a consumer framing sweep —
per `docs/hypotheses.md`, not a silent claim).
