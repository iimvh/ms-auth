# ms-auth

An OpenAI-compatible HTTP proxy for **Microsoft 365 Copilot**, with two authentication paths: `msal` for work/school accounts and `personal` for consumer Microsoft accounts.

Point an OpenAI-compatible client at `http://127.0.0.1:4141/v1`. The proxy authenticates to M365, opens the Copilot WebSocket, translates messages to and from the OpenAI format, streams replies, and turns fenced Markdown into OpenAI `tool_calls`.

> Unofficial and reverse-engineered from the M365 web client. Not affiliated with, endorsed by, or supported by Microsoft. The protocol can change without notice, and using it may conflict with Microsoft's terms. Use it only with an account you are allowed to use this way.

## Features

- **OpenAI Chat Completions API** — `POST /v1/chat/completions`, JSON and SSE streaming, `stream_options.include_usage`, text and image inputs, tool definitions and `tool_choice`.
- **Two auth modes** — enterprise MSAL login (PKCE + TOTP, automated or interactive) and personal-account capture/replay over Chrome DevTools Protocol.
- **Tool calling** — M365 has no native OpenAI tool API, so the proxy injects fenced tool instructions and parses the model's fenced blocks back into OpenAI `tool_calls`. A shell-first framing path is what makes real multi-turn agent loops work.
- **Model routing** — maps OpenAI-style model names to M365 Copilot tones, including Claude Sonnet/Opus and the GPT-5.x/6.x family. See `GET /v1/models`.
- **Multi-account personal mode** — a registry of accounts, each with its own Chrome profile, CDP port and session record, switchable per request. Includes a local dashboard at `/accounts/ui`.
- **Throttle management** — optional turn-budget pacing (`M365_AVOID_THROTTLING=1`), degradation backoff, priority-access tracking and quota diagnostics in `usage.x_m365_*`.
- **Frame replay harness** — recorded M365 frames can be replayed through the real transport without spending a turn.
- **Health endpoints** — liveness plus credential readiness for the active account.

## How it fits together

```
OpenAI-compatible client ──HTTP──▶  ms-auth :4141  ──WebSocket──▶  Microsoft 365 Copilot
Claude Code ──Anthropic Messages──▶ ms-gateway :4142 ──OpenAI──▶  ms-auth
```

`ms-gateway` is a separate Python project that translates Anthropic Messages to the OpenAI API. It is not part of this repository; this repository is the proxy and the M365 protocol implementation.

## Repository layout

| Package | Purpose |
|---|---|
| `@m365-copilot/core` | Auth adapters, M365 transport and frame types, consumer chathub code, framing, tool parsing, routing, budgets, image I/O, account registry and the `m365-personal` CLI. |
| `@m365-copilot/proxy-lib` | OpenAI request schemas, conversation/session pool, request handler, health payload and accounts API. |
| `@m365-copilot/proxy` | Nitro HTTP server: routes, startup auth plugin, CORS and the accounts dashboard. |

The source follows ports and adapters: `packages/core/src/turn-context.ts` defines the transport-facing contract, `adapters/msal` and `adapters/personal` hold the two credential implementations, and `protocol/` never learns which one produced a turn. See [ADR-0001](docs/adr/0001-ports-and-adapters.md).

## Requirements

- Node.js 24 or newer, pnpm 10 or newer.
- A Microsoft 365 Copilot account (work/school or personal).
- Chrome/Chromium. Enterprise login drives a browser through Playwright; personal mode attaches to your own Chrome over CDP.
- A windowed browser for personal mode. Headless Chrome cannot complete the sign-in interaction.

## Quick start

```sh
pnpm install
pnpm build
```

Pick one auth path.

### Enterprise (`msal`)

Create `~/.config/opencode-m365/secrets.json` with the TOTP **seed**, not a six-digit code:

```json
{
  "email": "you@company.com",
  "password": "your-password",
  "mfaSecret": "JBSWY3DPEHPK3PXP"
}
```

If the tenant uses push-only MFA, skip the secrets file and approve the sign-in by hand:

```sh
M365_ENABLE_INTERACTIVE_APPROVAL=1 pnpm dev
```

### Personal (`personal`)

The consumer sign-in is behind Cloudflare, so the proxy cannot launch the browser itself. Start your own Chrome on a dedicated profile with a debugging port, then attach:

```sh
chrome --remote-debugging-port=9222 \
  --user-data-dir="$HOME/.config/opencode-m365/personal-profile" \
  https://m365.cloud.microsoft/chat

M365_PERSONAL_CDP_URL=http://127.0.0.1:9222 pnpm m365-personal login
pnpm m365-personal status
```

Keep that Chrome window open with the chat tab loaded. The session is re-minted by attaching to it, so closing the window eventually breaks the account until you log in again.

### Run it

```sh
# enterprise
M365_AVOID_THROTTLING=1 pnpm dev

# personal
M365_AUTH=personal M365_AVOID_THROTTLING=1 pnpm dev
```

PowerShell equivalents:

```powershell
$env:M365_AUTH = "personal"
$env:M365_AVOID_THROTTLING = "1"
pnpm dev
```

Verify:

```sh
curl http://127.0.0.1:4141/v1/models
curl http://127.0.0.1:4141/healthz
```

## API

| Endpoint | Description |
|---|---|
| `GET /health` | Liveness: `{"status":"ok"}`. |
| `GET /healthz` | Credential and account health. Never mints and never spends a turn. Always HTTP 200; the body carries the verdict. |
| `GET /v1/models` | Models this proxy accepts, with advertised context and output limits. |
| `POST /v1/chat/completions` | OpenAI-compatible chat completion. `stream: true` returns SSE. |
| `GET /accounts` | Personal mode only: accounts, health and registry location. |
| `POST /accounts` | Add an account (`{ "id", "label" }`) and receive its Chrome/login commands. |
| `PATCH /accounts/:id` | Rename an account's label. |
| `DELETE /accounts/:id` | Remove an account; refuses with `409 account_busy` while it serves a conversation. |
| `POST /accounts/:id/activate` | Switch the active account for subsequent requests. |
| `GET /accounts/ui` | Local accounts dashboard. |

### Chat completion

```sh
curl http://127.0.0.1:4141/v1/chat/completions \
  -H 'Content-Type: application/json' \
  -d '{
    "model": "claude-sonnet",
    "messages": [{ "role": "user", "content": "Explain what this project does in one sentence." }]
  }'
```

Add `"stream": true` for SSE. With `tools`, a tool-using turn returns `finish_reason: "tool_calls"` and standard `tool_calls` entries; send each result back as a `role: "tool"` message with the matching `tool_call_id`. By default only the first tool call in a turn is executed; set `M365_ALLOW_MULTI_TOOL=1` to allow more.

Responses include M365 diagnostics in the `usage` object (`x_m365_conversation_pct`, `x_m365_turn_count`, `x_m365_dea_score`, priority-access remaining values, and so on).

### Model selection

| Requested model | Routes to | Notes |
|---|---|---|
| `claude-sonnet`, `claude-sonnet-4.5`, `claude-sonnet-4.6` | `Claude_Sonnet` | Included scenario on most accounts; the usual default for coding loops. |
| `claude-sonnet-5.5` | `Claude_Sonnet` | Paid scenario with its own priority-access budget. |
| `claude-sonnet-5` | `Claude_Sonnet_5` | Paid scenario. |
| `claude-opus-4.8`, `claude-opus-4.5` | `Claude_Opus` | Included scenario. |
| `claude-opus`, `claude-opus-5`, `claude-opus-5.5` | `Claude_Opus` | Paid scenario and subject to priority-access limits. |
| `gpt-5.5-think-deeper` | `Gpt_5_5_Reasoning` | Default when `model` is omitted in enterprise mode. |
| `gpt-5.5`, `gpt-5.6`, `gpt-5.4`, `gpt-5.3`, `gpt-5.2`, `gpt-6-*` | Matching GPT tones | Availability depends on the account and tenant. |
| `m365-copilot`, `auto` | `magic` | M365's default Copilot tone. |

Personal mode always answers with the consumer Copilot model (`tone: "Magic"`); requested Claude/GPT tones are not served. Its default model is the one recorded by the login capability probe.

## Configuration

All settings are environment variables. The data directory defaults to `~/.config/opencode-m365`.

### Runtime

| Variable | Default | Meaning |
|---|---|---|
| `M365_AUTH` | `msal` | `msal` or `personal`. Anything else fails at boot. |
| `HOST` / `NITRO_HOST` | unset | Bind address. Set `127.0.0.1` explicitly to keep the proxy local. |
| `PORT` / `NITRO_PORT` | `4141` | Listen port. `m365-proxy [port]` also accepts a port. |
| `M365_ACCOUNT_HOME` | `~/.config/opencode-m365` | Base directory for the account registry, account sessions and personal profiles. |
| `M365_AVOID_THROTTLING` | off | Pace turns against the local budget instead of tripping `PerUserThrottled`. |
| `M365_BUDGET_CAPACITY` | `100` | Turn-budget bucket capacity. |
| `M365_BUDGET_REFILL` | `1.6` | Tokens refilled per minute. |
| `M365_BUDGET_RESERVE` | `20` | Reserve held below the bucket floor. |
| `M365_BUDGET_HOLD_MIN` | `75` | Hold duration after a throttle, in minutes. |
| `M365_BUDGET_MAX_WAIT_MIN` | `15` | Maximum wait for a turn slot. |
| `M365_SAVE_HISTORY` | off | Temporary chat by default; `1` saves conversations to M365 history. |
| `M365_CONTEXT_WINDOW` | `1000000` | Context window advertised in `/v1/models`. |
| `M365_MAX_OUTPUT_TOKENS` | `1000000` | Max output advertised in `/v1/models`. |
| `M365_SESSION_IDLE_MS` | `3600000` | Idle time before a pooled conversation is forgotten. |
| `M365_OUTPUT_CHAR_CEILING` | `12000` | Character count that yields `finish_reason: "length"`; `0` disables. |
| `M365_EARLY_STOP_AFTER` | `2000` | Early-stop threshold when tools are present. |
| `M365_ALLOW_MULTI_TOOL` | off | Allow more than the first tool call. |
| `M365_CONFAB_RETRIES` | `1` | Retries after the model fabricates a tool result. |
| `M365_FORCE_AGENT` | unset | `1` always uses the Copilot Studio agent for tool requests; `0` never does. |
| `M365_NO_IMAGE_GEN` | unset | Disable generated images. |
| `M365_SCENARIO` / `M365_LICENSE_TYPE` | per model | Override the upstream scenario and license strings. |
| `M365_FRAMING_VARIANT` / `M365_FRAMING_FILE` | per tone | Force a tool-framing variant. |

### Enterprise auth

| Variable | Default | Meaning |
|---|---|---|
| `M365_SECRETS_FILE` | `~/.config/opencode-m365/secrets.json` | Email, password and TOTP seed. |
| `M365_CACHE_FILE` | `~/.config/opencode-m365/msal-cache.json` | MSAL token cache. |
| `M365_BROWSER_PROFILE` | `~/.config/opencode-m365/browser-profile` | Playwright profile used during login. |
| `M365_ENABLE_INTERACTIVE_APPROVAL` | unset | `1` allows manual sign-in in a visible browser. |
| `M365_NO_INTERACTIVE` | unset | `1` forbids interactive login. |
| `M365_INTERACTIVE_TIMEOUT_MS` | `600000` | Manual sign-in timeout. |
| `M365_LOGIN_LOCALE` | `en-GB` | Login page locale. |
| `M365_LOGIN_TIMEZONE` | `Europe/Copenhagen` | Login page timezone. |
| `M365_LOGIN_UA` | Chrome UA | User agent used during login. |
| `CHROMIUM_PATH` | autodetected | Browser binary. |

### Personal auth

| Variable | Default | Meaning |
|---|---|---|
| `M365_PERSONAL_CDP_URL` | unset | Chrome CDP endpoint to attach to. Required to pass Cloudflare. |
| `M365_PERSONAL_MODEL` | probed model | Default upstream model in personal mode. |
| `M365_PERSONAL_MINT_STRATEGY` | `harvest` | `harvest` attaches to Chrome; `capture` launches a browser. |
| `M365_PERSONAL_MINT_MIN_INTERVAL_MS` | `600000` | Minimum interval between mints. |
| `M365_PERSONAL_MINT_MAX_PER_HOUR` | `3` | Mint budget per hour. |
| `M365_PERSONAL_TOKEN_TTL_MS` | `3600000` | Estimated token lifetime; a staleness policy, not a measurement. |
| `M365_PERSONAL_PROBE` / `M365_PERSONAL_PROBE_MODELS` | `1` / `claude-sonnet` | Login-time capability probe. |
| `M365_PERSONAL_CODE_INTERPRETER` | off | `1` re-enables M365's Python sandbox in personal mode. |
| `M365_PERSONAL_PROFILE` / `M365_PERSONAL_SESSION_FILE` | unset | Legacy path overrides that beat the account registry for every account. Leave unset in a multi-account setup. |

The complete personal-mode matrix, including login and harvest timing knobs, is in [docs/personal-auth.md](docs/personal-auth.md).

### Diagnostics

| Variable | Default | Meaning |
|---|---|---|
| `M365_DEBUG` | off | Write the debug log. |
| `M365_TRACE` | off | Full, untruncated logging. |
| `M365_LOG_STDOUT` | off | Also write logs to stdout. |
| `M365_LOG_FILE` | `~/.config/opencode-m365/debug.log` | Log path. |
| `M365_DUMP_FRAMES` | off | Write NDJSON frame dumps for replay. |
| `M365_FRAME_DIR` | `~/.config/opencode-m365/frames` | Frame dump directory. |

## Personal mode limits

Personal mode is a capture-and-replay implementation, not an OAuth client. It has real constraints:

- A free consumer account gets **80 turns per day** (`CostQuota`, reset at 00:00 UTC). Every probe, retry and mint counts.
- The consumer endpoint serves its own Copilot model only. Claude/GPT tones return `InternalError`.
- No Copilot Studio agent, no image generation, and the Python sandbox is off by default.
- Re-minting requires a running, windowed Chrome with a `m365.cloud.microsoft/chat` tab. There is no headless path.
- The consumer token is an opaque JWE with an **estimated** expiry, renewed by harvesting a fresh token from the page socket.
- A conversation belongs to the account it was opened on. Switching accounts affects the next request, not a thread in flight.
- Run one proxy per data directory. Two processes sharing an account can both believe they own its mint lock.

Use enterprise mode for long agent sessions. Personal mode is suitable for demos and light chat.

## Development

| Command | Purpose |
|---|---|
| `pnpm dev` | Nitro dev server on `:4141`. |
| `pnpm build` | Build `core`, `proxy-lib` and the proxy. |
| `pnpm proxy` | Run the built Nitro server. |
| `pnpm typecheck` | TypeScript check. |
| `pnpm test` | Build and run the unit suite. |
| `pnpm test:live` | Opt-in live tests against M365; spends real turns. |
| `pnpm lint:boundaries` | Enforce the adapter boundary rule from ADR-0001. |
| `pnpm snapshot:exports:check` | Verify the public `core` export surface has not changed. |
| `pnpm snapshot:exports` | Update that snapshot deliberately. |
| `pnpm m365-personal login` / `status` / `logout` | Manage personal sessions. |

The replay harness removes M365 from the test loop: `M365_DUMP_FRAMES=1` writes recorded frames, and the test suite replays them through the real transport. See [docs/replay-harness.md](docs/replay-harness.md).

Nix users can run `nix develop`, `nix run .`, or import the NixOS module `services.m365-copilot-proxy`.

## Documentation

- [docs/run-guide.md](docs/run-guide.md) — full stack: proxy, Anthropic gateway, Claude Code.
- [docs/personal-auth.md](docs/personal-auth.md) — consumer auth, accounts, minting, recovery.
- [docs/m365-copilot-api.md](docs/m365-copilot-api.md) — protocol and endpoint notes.
- [docs/tool-calling.md](docs/tool-calling.md) — fenced tool-call contract and enforcement layers.
- [docs/prompt-engineering.md](docs/prompt-engineering.md) — framing variants and measured results.
- [docs/hypotheses.md](docs/hypotheses.md) — measurement log; source of truth for protocol findings.
- [docs/replay-harness.md](docs/replay-harness.md) — frame fixtures and offline replay.
- [docs/adr](docs/adr) — architecture decisions.
- [GLOSSARY.md](GLOSSARY.md) — canonical project vocabulary.

## Security

- The proxy performs **no API-key validation** and has permissive CORS. Keep it bound to loopback and never expose it to an untrusted network.
- `~/.config/opencode-m365` holds live bearer tokens and, for personal mode, a browser profile with a full Microsoft account session. Treat it as a credential store.
- `pnpm m365-personal logout` deletes both the stored session and the browser profile. Use it on shared machines.
- Debug logs and frame dumps can contain prompts and model output. Review before sharing.

## License

MIT — see [LICENSE](LICENSE).
