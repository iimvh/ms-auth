# Architecture handoff — ms-auth / ms-gateway (Oct 22 2026)

State at handoff, decisions made, and the module plan agreed for the next
session. Nothing here is implemented yet; the code is exactly as shipped in
`ms-auth@a91f2d7` (2 commits) and `ms-gateway@df916ef`.

## Where the code is

| Repo | Branch | State |
|---|---|---|
| `iimvh/ms-auth` (this repo) | `master` | M365 proxy + personal auth. 1 squashed feature commit + 1 docs commit. Working tree clean. |
| `iimvh/ms-gateway` | `main` | Anthropic→OpenAI gateway (Python, 1.5k LOC, 98 tests green). |

Live status: personal (consumer) auth works end to end — `login` (Chrome over
CDP) → capability probe → `/v1/chat/completions` → gateway `/v1/messages`.
Blocked only by the free account's `CostQuota` (80 turns/day, resets 00:00 UTC).
See `docs/hypotheses.md` F97–F109 for every consumer finding.

## Current layout (technical layers, not domains)

```
packages/core/src/
  auth.ts               721  enterprise MSAL (PKCE, TOTP, automated/interactive)
  personal-*.ts        1467  personal capture/replay: session, capture, auth, cli, mode
  session.ts           1188  WS transport + protocol + personal turn path (mixed)
  fenced.ts            1475  framing prompt strings + fenced parser (mixed)
  model.ts              210  agent routing + retry
  agent.ts              402  Copilot Studio agent
  turn-budget.ts        433  throttle math
  priority-access.ts    211  metered budgets
  tools.ts / copilot.ts / image.ts
packages/proxy-lib/src/
  handler.ts           1231  OpenAI ⇄ M365 pipeline (state machine, implicit)
packages/proxy/               Nitro HTTP server (routes/, plugins/)
```

The split today is by *package*, not by *domain*. `core` holds eight domains and
nothing enforces boundaries: `session.ts` imports `personal-auth`,
`personal-capture`, `turn-budget`, `agent`, … freely.

## What the two projects actually are

- **`ms-auth` (TS, keep TS).** 90% of the work is reverse-engineering an
  undocumented protocol that Microsoft changes. What matters is the loop
  "frame looks different → know what to edit": same-language frame dumps,
  `zod.safeParse` for new frame shapes, Playwright/CDP (porting the capture
  driver to another language means rewriting 540 lines of the hardest code),
  sub-second tests. Go/Rust would raise the cost of exactly the work that
  happens most often; Python is equal on speed but loses frame type-safety
  (the `tone`-in-the-wrong-function bug from Oct 22 is that failure mode).
- **`ms-gateway` (Python, keep Python for now).** Stable, stateless translator,
  ~1.5k LOC, 98 green tests. It is not a performance bottleneck: one client,
  a few requests a minute, each dominated by M365 latency. Go would buy a
  single binary (~15 MB vs Python runtime) and less RAM, nothing else. Only
  worth it if distributing binaries to other people; then write a small Go
  gateway alongside and diff it against Python, don't big-bang rewrite.
  Rust: no.

Stack verdict: **no stack change**. The leverage is architecture and
fixture-based testing, not the language.

## Target architecture: modulith by domain

One binary, hard module boundaries, one-way dependencies.

```
src/
  modules/
    auth/                 public: authFor(mode): AuthProvider
      work/               MSAL, PKCE, TOTP, automated + interactive login
      personal/           capture, replay, mint, CLI
    protocol/             public: openTurn(ctx) / TurnStream / frame schemas
      transport.ts        WS connect, handshake, frame split, deflate
      frames.ts           zod schemas for M365 frames
      payload.ts          build turn arguments (tone, clientInfo, optionsSets)
      stream.ts           parse type:1/2/3 → text, early-stop
    framing/              public: framing(name), parseFenced()
      variants/           one object per framing, no if-chains
      parser/             fenced tool-call parser
    agent/                public: getOrCreateAgent()
    budget/               public: awaitTurnBudget(), meteredBudgetOf()
    http/                 public: createApp()   (today's proxy-lib)
    image/                public: generateImage()
  app/                    composition root
    server.ts             Nitro entry
    cli.ts                m365-personal
```

**The load-bearing rule:** `protocol/` must not import `auth/`. Auth supplies a
context object and nothing else:

```ts
interface TurnContext {
  token: string;
  wsUrl(ids: { requestId: string; sessionId: string; conversationId: string }): string;
  headers(conversationId: string): Record<string, string>;
  payloadDefaults(): Partial<TurnArguments>;
  capabilities: {
    kind: "enterprise" | "consumer";
    tone: string | null;      // "Magic" for consumer, model tone otherwise
    toolAgent: boolean;
    imageGeneration: boolean;
    codeInterpreter: boolean;
  };
}
```

Allowed imports:

```
http      → protocol, auth, framing, budget, image
app       → everything
auth      → msal / playwright / otpauth only
protocol  → zod, ws only
framing   → nothing (pure strings + parser)
agent     → auth (for its token), protocol types
budget    → fs only
```

Enforce with `dependency-cruiser` (or `eslint-plugin-boundaries`) in CI: a single
rule "modules/protocol must not import modules/auth" is the whole point.

Why domain rather than technical layer: layers change together (one Microsoft
frame change touches frames + stream + payload + sometimes handler). Domains
change for different reasons. Every bug fixed on Oct 22 lived in exactly two
domains — `protocol` and `auth/personal`.

## Migration order (safe, incremental — no big bang)

1. **Add dependency rules for the *current* layout** (`dependency-cruiser`),
   record the baseline, see the real coupling. Cheap, no behaviour change.
2. **Golden-frame fixtures.** Turn existing `M365_DUMP_FRAMES=1` dumps into
   `test/fixtures/frames/*.ndjson` and add a test that replays them through the
   transport/stream code. This is the prerequisite for everything below: it is
   the only way to prove a refactor did not change protocol behaviour, and it
   does not cost M365 turns.
3. **Extract `auth/`** out of `core` (work + personal + the mode switch). Add a
   test asserting `protocol` does not import `auth`; it fails first at
   `session.ts`, which is the signal to do step 4.
4. **Extract `protocol/`**: split `session.ts` into transport / frames / payload /
   stream; replace every auth import with the injected `TurnContext`.
   Hardest step — do it with step 2's fixtures green.
5. **Split `framing/`** out of `fenced.ts`: each variant becomes a data object
   (`{name, render(tools), transcriptStyle}`) instead of scattered branches.
6. **Collapse the package layout** into `src/modules/*` + `src/app/*`, still one
   binary (`m365-proxy`) and one CLI (`m365-personal`).

## Context window and handoff (agreed, not implemented)

The advertised window is wrong today: `proxy-lib/src/index.ts` sends
`context_window: 1_000_000` for *every* model, which is a Gemini-tier number
Copilot does not use. Working assumption until measured:

| Mode | `M365_CONTEXT_WINDOW` |
|---|---|
| personal (consumer `auto`) | **200_000** (chosen; conservative middle of 128k–200k) |
| enterprise default | 200_000 |
| premium Opus / GPT-6 | 200_000 (1M is a separate tier, unproven here) |

`M365_MAX_OUTPUT_TOKENS` should be 32_000, not 1M. Context should become
per-model in `/v1/models`, not one constant.

Two follow-ups, in order:

1. **Do not open a new conversation while a session is under the ceiling.**
   Audit result (Oct 22): there is no such spam today — the six
   `newConversation()` calls in `handler.ts` fire only on a non-continuation
   re-send (227, measured: reusing contaminates, 7/10 → 3/10), a model/agent
   fallback (380, 622, 653, 728) or a Disengaged retry (687, a Disengaged
   conversation stays Disengaged). The real gap is that nothing opens a new
   conversation *because the context is nearly full* — that is item 2. Worth
   reviewing separately: `SessionPool` evicts an idle session at 30 minutes,
   which abandons a still-live M365 conversation (costs no turn, but ends a
   long session's continuity).
2. **Measure tokens, then hand off.**
   - `usage.x_m365_prompt_tokens_est`: estimate the prompt actually sent
     (`messages` + `tools` + framed text, chars/4 or a real tokenizer). Offline,
     no M365 turn, and the only honest basis for any threshold. Note the
     existing `x_m365_conversation_pct` is *messages*/600, not tokens — useless
     as a context gauge.
   - Thresholds: ≤60% normal; 60–80% warn (surface the real percentage to the
     client); >80% proactively `newConversation()`.
   - Handoff is cheap here: the harness re-sends the whole history every
     request and `formatMessages` re-renders it, so a fresh M365 conversation
     loses nothing; an optional one-time summary turn is the only cost.
   - Optional, log-only heuristic: the model asking again for something already
     in the history (a created file, a read path) means M365 truncated the
     context underneath us.

### Limits differ by mode — handoff thresholds must too

| Tier | Personal (free consumer) | Enterprise |
|---|---|---|
| Daily turn cap | **`CostQuota` 80/day** (reset 00:00 UTC), all models | none for ordinary models |
| Per-conversation | 600 messages | 600 messages |
| Thread rate | applies | **`PerUserThrottled`** — ~100 turns, refill ~1.6/min, ~1 h hold once tripped |
| Per-model priority | none | Opus 5.5 40/day + 75/week; Sonnet 5.5 80/day + 150/week (only these) |

Consequence: a handoff opens a new M365 conversation, and *conversations* are
what the enterprise thread-rate throttle counts. So the threshold should differ
by mode — hand off earlier on personal (a new thread there costs nothing beyond
the turn, and the turn is already counted), later on enterprise
(~85–90%, so a long pi session keeps one thread). Recorded here, not built.

## Smaller follow-ups agreed (not architecture)

- **Gateway observability** (its weakest point): one INFO line per request —
  `model_in → model_upstream, messages=N, tools=? → status, ms, x-m365 quota
  left`. Today only `log.debug("upstream completion in %d ms")` exists, so a
  Claude Code failure cannot be attributed to Disengaged vs 429 vs a hung
  upstream. ~30 lines, no logic change.
- Gateway confirmed sound on timeouts (600 s turn, 15 s probes) and on
  deliberate non-retry (error mapping preserves 400/429/503/502 for the client
  to act on; retrying inside the gateway would burn personal quota).
- Personal mode: free tier is 80 turns/day; never bench against it.

## Open questions for next session

- Do we keep `M365_AUTH` as the only switch, or add a third mode later (Entra
  B2C, another tenant type)? The `AuthProvider` interface above is shaped so
  that adding one does not touch `protocol/`.
- `framing/` variants are prompt engineering artifacts that change with bench
  results — decide whether they live in code or in versioned data files.
- The gateway may eventually want to merge into the same repo as `http/`
  consumers (the Anthropic surface is just another adapter over the same
  `openTurn`); not now.
