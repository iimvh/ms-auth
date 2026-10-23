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
  expiry has to be **estimated** (`M365_PERSONAL_TOKEN_TTL_MS`, default 1 h).
  That 1 h is a default, **not a measurement**: a captured token kept serving
  turns **at least 77 minutes** after capture (measured on a free consumer
  account, 2026-10-23), and the MSAL credential the page derives it from lives
  **8 hours** (`cachedAt`→`expiresOn` in the page's own MSAL cache). The real
  lifetime of the socket token is **unknown**. Treat the estimate as a
  staleness policy, not as a fact — and never as "(re)log in every hour".
- At runtime the proxy rebuilds the socket URL from those captured values, so
  a Microsoft param change is fixed by re-running one command — no table to edit.
- When the token is near expiry it is re-minted by **harvesting**: the proxy
  attaches to your running Chrome over CDP, sends one short message to make the
  page open its Chathub socket, and reads the token off that socket. This is the
  default (`M365_PERSONAL_MINT_STRATEGY=harvest`) and takes ~2 s. Launching a
  Playwright browser (`=capture`, the old behaviour) is still available but
  cannot pass Cloudflare itself (F101), so it is only a deliberate opt-out.
- **Measured caveats about re-minting** (2026-10-23, free consumer account):
  - The page opens its Chathub socket **only when a message is sent** — not on
    page load, not when the composer is focused, and it closes the socket ~6 s
    after opening. There is no long-lived connection to piggyback on.
  - A send **usually** yields a new token — 5 of 6 rounds, 10 min apart, each
    got a different one. It is not guaranteed: two sends 3 minutes apart once
    returned the *same* token. A harvest reports that (`rotated: false`) rather
    than pretending a new credential was obtained.
  - The socket token exists **nowhere in page storage** (`localStorage`,
    `sessionStorage`, IndexedDB, Cache Storage). The only way to obtain one is
    to provoke a connection. The page's MSAL cache holds tokens for *other*
    scopes, not this one.
  - There is **no refresh token exposed to the proxy**. The page's MSAL cache
    did carry a `refreshtoken` entry once and it then disappeared, so it is not
    a stable contract to build on. Renewal happens inside the page, not in
    `personal-capture.ts`.
  - **A harvest costs one turn** even when the token did not rotate, because the
    turn is spent on the *send*. The mint budget counts sends, not rotations.
  - It needs **a chat tab open in that Chrome**. Chrome running with the
    debugging port but no `m365.cloud.microsoft` tab is `no_chat_tab`, not a
    silent failure.

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
| `M365_PERSONAL_MINT_MIN_INTERVAL_MS` | `600000` | Mints are at least this far apart. |
| `M365_PERSONAL_MINT_MAX_PER_HOUR` | `3` | Mint budget per hour (each mint is a fresh conversation — the thread-rate throttle's unit). |
| `M365_PERSONAL_MINT_STRATEGY` | `harvest` | `harvest` attaches to your Chrome and sends one message; `capture` launches a browser (the old behaviour; cannot pass Cloudflare on its own). |
| `M365_PERSONAL_MINT_PROMPT` | `.` | The message a harvest sends to force the socket. Costs one turn. |
| `M365_PERSONAL_MINT_SOCKET_WAIT_MS` | `20000` | How long a harvest waits for the Chathub socket after sending. |
| `M365_PERSONAL_MINT_WAIT_MS` | `10000` | How long a mint waits for the app's socket on page load before typing anything. |
| `M365_PERSONAL_LOGIN_WAIT_MS` | `300000` | How long `login` (headed) waits on the Microsoft sign-in page for you to finish signing in. |
| `M365_PERSONAL_CDP_URL` | unset | Attach to an already-running Chrome (`http://127.0.0.1:9222`) instead of launching one. Required to clear Cloudflare. |
| `M365_PERSONAL_AUTOMATED_LOGIN` | `1` | `0` = never drive the sign-in form with `secrets.json`; wait for a human. |
| `M365_PERSONAL_TOKEN_TTL_MS` | `3600000` | Estimated lifetime of the opaque JWE token. **A guess, not a measurement** — one token served ≥ 77 min; see the note under "What this path is — and is not". |
| `M365_PERSONAL_VERIFY_WAIT_MS` | `300000` | How long a capture waits for a human to clear the verification/security-check dialog, then sends again. Nothing bypasses the challenge. |
| `M365_PERSONAL_CODE_INTERPRETER` | unset | `1` re-enables M365's Python sandbox optionsSets in personal mode. |
| `M365_PERSONAL_PROFILE` | `<data dir>/personal-profile` | Browser profile location. |
| `M365_PERSONAL_SESSION_FILE` | `<data dir>/personal-session.json` | Record location. |
| `M365_ACCOUNTS_FILE` | `<data dir>/accounts.json` | The account registry (see "Several accounts"). Governs personal mode only; MSAL mode ignores it. |
| `M365_ACCOUNT_HOME` | `$HOME/.config/opencode-m365` | The data dir itself — overrides `accountsDir()` and the layout defaults, for tests and side-by-side installs. |

`M365_PERSONAL_PROFILE` and `M365_PERSONAL_SESSION_FILE` are **overrides**, not
the default paths any more: when either is set it wins over the registry for
every account, which is almost never what a multi-account setup wants (see the
warning in "Several accounts").

## Several accounts

One data dir can hold several personal accounts and serve one at a time. An
account is a **pair**:

| Part | Per account | Holds |
|---|---|---|
| session record | `accounts/<id>/session.json` | the minted token + captured params |
| browser profile | `accounts/<id>/profile/` | that account's Microsoft session cookies |
| debugging port | `accounts.json` | the CDP port that profile's Chrome listens on |

`accounts.json` is the registry (`version: 1`):

```json
{
  "version": 1,
  "activeId": "work",
  "accounts": [
    { "id": "default", "label": "Personal (free)", "cdpPort": 9222,
      "sessionFile": "…/accounts/default/session.json",
      "profileDir": "…/accounts/default/profile",
      "createdAt": "2026-10-23T05:00:00.000Z" }
  ]
}
```

Paths are stored absolute, so a registry keeps working from a different working
directory. The account that existed before this feature (`personal-session.json`
+ `personal-profile/` directly in the data dir) becomes the `default` entry in
place — its files are **not moved**, the registry just points at them. So a
single-account install keeps working with no migration step, and `default` keeps
the old paths forever (which is why `default` is worth keeping as a name).

**Add an account** (the proxy has to be running with `M365_AUTH=personal`):

1. Open `http://127.0.0.1:<port>/accounts/ui` (loopback only — the page has no
   auth of its own, so never expose the proxy beyond your machine).
2. **Add account**: id (`a-z0-9-`, up to 32 chars) and a label. The page shows
   two commands.
3. Run the **Chrome** one, sign in to the account, and leave that window open —
   `chrome --remote-debugging-port=9223 --user-data-dir="…/accounts/work/profile" https://m365.cloud.microsoft/chat`
4. Run the **login** one: `pnpm m365-personal login --account work`.
5. `pnpm m365-personal status --account work` shows its expiry, scenario and
   profile state; `/accounts/ui` shows every account's health.
6. Tick the account's radio in the UI (or `POST /accounts/work/activate`) to
   switch. The switch applies to the **next request**; a thread already running
   keeps the account it was opened on.

**One Chrome per account, one port per profile.** Chrome's debugging port is
per browser process, not per profile: two Chromes opened with
`--user-data-dir` for two accounts both take 9222 and the second is silently
just another tab of the first. So each account gets its own port, allocated at
creation from 9222 upward, and the proxy derives every account's CDP URL from
its own `cdpPort`. Keep one keeper per account:

```powershell
pwsh scripts/chrome-keeper.ps1 -Account work
```

`-Account` reads that account's `cdpPort`/`profileDir` from `accounts.json`;
`-Port`/`-ProfileDir` still override it, and without `-Account` the keeper keeps
the old single-account defaults.

**The switch is per request, not per process.** The mint mutex, the credential
cache, the degradation backoff, the turn-budget bucket and the conversation
fingerprint are all keyed by account — a throttled or exhausted account slows
down nothing else. `/healthz` reports the **active** account's credential state
plus `accounts: { count, activeId }`; it never probes every account (the accounts
page does, and caches the probes for 10 s).

**A thread belongs to the account it was opened on.** A conversation keeps its
account for its lifetime, so switching accounts mid-thread does not re-route it;
`DELETE /accounts/<id>` refuses with `account_busy` while any conversation of
that account is still live.

**The override trap.** `M365_PERSONAL_PROFILE` and `M365_PERSONAL_SESSION_FILE`
predate the registry and now beat it: with either set, every account resolves to
that one path. That was the only behaviour possible before multi-account, but in
a registry it means both accounts share one record and one profile — the second
login overwrites the first and both mints land on one identity. Leave both unset
when the registry is in use; the proxy warns when a registry exists and an
override is set.

**One data dir per operator per machine.** Two proxies with different
`M365_AUTH` or `M365_ACCOUNT_HOME` are independent; two proxies sharing a data
dir share the registry and the accounts (that is usually what you want), but
each account's Chrome can only be driven by one proxy at a time — the mint lock
is per process, so two processes can both think they own the account. Run one
proxy per data dir.

**What the page cannot do.** It is plain HTML served by the proxy: it cannot
start Chrome, cannot sign you in, and cannot see how many turns an account has
left today (M365 does not tell us). A green row means "record fresh and Chrome
answered". When an account's turns run out, it fails alone — switch manually.

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
- **`could not reach Chrome on http://127.0.0.1:9222 … start it on the capture
  profile: chrome --remote-debugging-port=9222 …`** (`chrome_unreachable`) —
  nothing is listening on the CDP port, so the harvest had nothing to attach to.
  Start Chrome as quoted. This used to advise `login`, which cannot help: login
  does not open a debugging port.
- **`Chrome on … has no m365.cloud.microsoft/chat tab open`** (`no_chat_tab`) —
  Chrome is reachable but the chat page is not among its tabs. Open it (or
  `login`).
- **`a message was sent but the page opened no chat socket`** (`mint_not_rotated`)
  — the send did not produce a Chathub socket (a gate dialog, a broken page).
  Retry, or `login`.
- **`could not re-mint the M365 personal session — run: pnpm m365-personal login`**
  (`capture_failed`) — reachable, but no usable session could be read from it —
  transient network, the Cloudflare/security-check dialog, or a Microsoft
  param/flow change. Retry once; if it persists, repeat the CDP login.
- **`refusing to mint ... mint_rate_limited`** — the mint rate limits hit. Not
  dead: usually means you ran a fresh-conversation sweep against the personal
  account. Wait for the interval/hour budget, or `login` by hand.
- `M365_AUTH=<bad>` fails at boot: `M365_AUTH must be "msal" or "personal", got "<value>"`.

A dead session surfaces through the OpenAI API as
`401 {"error":{..., "code":"m365_personal_session_dead"}}` **before any stream
byte** is committed — a `stream:true` client gets a real HTTP status, not a
mid-stream error.

## Keeping Chrome there (the operating requirement, and its one limit)

A re-mint needs a Chrome running on the capture profile **with a chat tab open**.
`scripts/chrome-keeper.ps1` supervises that:

```powershell
pwsh scripts/chrome-keeper.ps1          # headed (the only mode that can mint)
pwsh scripts/chrome-keeper.ps1 -Once    # start it and return
```

It restarts Chrome if it dies, reopens the chat tab if it is closed, and says so
loudly (once) when the tab has landed on a sign-in page — which needs a human and
`pnpm m365-personal login`.

**`-Headed:$false` does not work.** Measured 2026-10-23: a headless Chrome starts,
loads the chat page, and still shows the signed-in session — then the composer
rejects every click, the Send button never appears, and the page reports a
verification/security check. The gate is on **interaction**, not on page load, so a
mint against headless fails with `mint_not_rotated`. The keeper warns if you try.
If a windowless host is ever required, the next step is a virtual display, not
headless mode.

## Health

`GET /healthz` reports whether this process can still serve a turn — the thing a
generic upstream 502 hides. It never mints and never spends a turn.

```json
{ "status": "ok", "authMode": "personal",
  "credential": { "status": "ok", "expiresInMs": 3594494,
                  "fresh": true, "cdpReachable": true } }
```

`credential.status`: `ok` (record fresh; a `hint` appears if Chrome is not
reachable, since that only matters once it expires), `stale` (past the estimated
expiry, Chrome reachable, the next request re-mints), `at_risk` (expired or
absent **and** no Chrome listening — the next request fails, and the `hint` quotes
the fix), `corrupt` (a record file **exists** but cannot be used — see below).
Always HTTP 200, verdict in the body. In MSAL mode:
`{"credential":{"applicable":false}}`.

### `mint_not_rotated` usually means the account is out of turns

The harvest sends one message to make the page open its Chathub socket. When the
free consumer account has spent its **80 turns/day**, the page **accepts the typing
but never opens a socket**, so the harvest waits out its timeout and reports
`mint_not_rotated` ("a message was sent but the page opened no chat socket").

That message is misleading: it reads like a browser problem, and the fix it
suggests (`login`) cannot help — the quota is per account, not per session, and a
fresh login does not reset it. What the page actually says, in the composer area:

> Bạn đã đạt đến giới hạn hàng ngày của mình. Nhận thêm mức sử dụng ngay hoặc
> quay lại lúc 7:00 SA

**How to tell them apart.** The composer is a `contenteditable` span, so
`contentEditable` stays `"true"` either way; the discriminator is Playwright's
**`isEnabled()`**, which is `false` while the page refuses input:

| page state | `isEnabled()` | sending works |
|---|---|---|
| idle | `true` | yes — socket opens in ~1 s |
| Copilot still streaming the previous probe | `false` | no |
| **daily quota exhausted** | `false` | no — and it stays that way until 00:00 UTC |

The harvest now **waits for the composer to be enabled** (up to
`M365_PERSONAL_COMPOSER_READY_MS`, default 45 s) before typing, and reports
`capture_failed` naming the busy/blocked page when it never is. Before that it
typed into a disabled editor, nothing was sent, and the operator got a
`mint_not_rotated` that pointed at the wrong thing.

If `/healthz` says `ok` and every mint still fails, check the quota before
touching Chrome: `m365-personal status` cannot see it, so look at the page.

### A broken record file is not "not logged in"

`PersonalStore.inspect()` separates four causes that used to collapse into one
`null`:

| `problem` | what it means | what to do |
|---|---|---|
| `absent` | no file — the ordinary first-run case | `pnpm m365-personal login` |
| `unreadable` | the file is there but cannot be read (permissions, a directory, a lock) | fix the path's permissions |
| `invalid_json` | the file is not JSON (truncated write, a bad hand-edit) | delete it, log in again |
| `wrong_shape` | valid JSON but not a record — the `hint` names the missing fields | delete it, log in again |

`pnpm m365-personal status` and `/healthz` report which of these it is, so
`wrong_shape` no longer looks like "you never logged in". A UTF-8 BOM (what
PowerShell's `Set-Content -Encoding UTF8` writes) is **stripped**, not treated as
corruption — that used to make a perfectly good record read as absent.

Recovery is automatic: a corrupt record does **not** need a manual delete before
a mint. The harvest needs a Chrome, not a stored record, so it runs anyway, writes
a good record over the bad one, and the proxy logs an `error` naming the problem.
Measured: `wrong_shape` → `/healthz` `corrupt` → mint in 2.0 s → record replaced →
`/healthz` `ok`.

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
5. force a stale record (edit expiresAt back 10 min) and run step 3 again
                                               → record: did it harvest? how long? (needs the
                                                 Chrome from step 1 still open ON A CHAT TAB;
                                                 expected ~2 s per F111. With no Chrome the
                                                 error must be chrome_unreachable, not a 502
                                                 after two failed launches)
6. rm -rf the personal profile, run step 3     → record: exact error text, no hang
7. after 24 h: pnpm m365-personal status --json → record: mint count, token lifetime
8. keep Chrome closed, force a stale record, run step 3
                                               → record: chrome_unreachable, and NO turn spent
```

Known numbers so far (F97–F111 in `docs/hypotheses.md`): a free consumer account
is `scenario=OfficeWebFreeConsumerCopilot`, `licenseType=Free`, **80 turns/day**
(`CostQuota`, reset 00:00 UTC); the socket token's real lifetime is **unknown**
(>= 77 min observed; the 1 h figure is a default, not a measurement — the page's
MSAL credential lives 8 h); a harvest mint took **2.3 s** end to end and wrote a
new token (2026-10-23); a captured session
answers `/v1/chat/completions` and an Anthropic gateway in front of it end to
end. Do not bench on a free account without counting turns.

Also worth recording when they run: whether the consumer scenario produces a
**fenced tool call** at all (spec §10.1: if not, phase 1 ends at "personal chat
works, tool routing unproven" and the next step is a consumer framing sweep —
per `docs/hypotheses.md`, not a silent claim).
