# 0001 — Ports and Adapters for `packages/core`

**Status:** accepted
**Supersedes:** `docs/architecture-handoff.md` §"Target architecture: modulith by domain" and §"Migration order"

## Context

`packages/core` holds ten responsibilities in one flat directory: MSAL enterprise
login, personal-account capture and harvest, the M365 WebSocket transport, framing
prompt strings, tool-call parsing, model routing, turn budgets, image generation,
Copilot Studio agent provisioning, and account selection. Nothing enforces a
boundary: `log.ts` is imported by 17 files and `personal-session.ts` by 9.

The real coupling problem is narrower than "domains are entangled". It is that

1. the transport (`session.ts`) decides the auth mode itself, so protocol code
   imports personal-account code (`session.ts:30–33`) and resolves accounts
   (`session.ts:447`, `session.ts:455`);
2. the choice of credential is made in two places, `auth.ts:686` and
   `model.ts:75`, and the two disagree in shape — one switches on
   `isPersonalMode()`, the other on whether an account is pinned;
3. the personal path talks to a different Microsoft endpoint (chathub) with a
   different URL, headers, payload and capability set, and that code is interleaved
   with the shared transport.

A modular monolith was designed for this in `architecture-handoff.md`: modules,
dependency tiers, four enforcement rules, four new packages. It was rejected
during design for three reasons.

- Every module would have exactly one implementation, so the interfaces exist to
  satisfy a rule rather than to allow substitution.
- Two of its own rules contradict the code: `auth/msal` must call personal
  minting (`auth.ts:686`), and `protocol` must resolve an account
  (`session.ts:447`). Honouring them required rewriting behaviour, not moving it.
- It changed the public API (`getToken` removed from the barrel, major bump) in a
  repo whose only consumers are three packages in the same workspace.

There is no test that protects the transport today: zero recorded frame fixtures
exist, and `proxy-lib/src/streaming.test.ts` mocks `ModelSession` wholesale, so it
stays green when the transport is broken.

## Decision

Structure `packages/core/src` as ports and adapters.

```
turn-context.ts    CredentialSource, SessionStore, TurnContext   (contracts only; ADR's port.ts)
adapters/msal/     token, login, secrets, artifact
adapters/personal/ mode, record, capture, harvest, harvest-cdp, auth, health, mint limits
protocol/          transport (session), frames (schemas, frame-replay), stream
personal-chathub/  URL, headers, params, payload and probe for the consumer endpoint
routing/           model and tone routing
budget/            turn budget, degradation backoff, priority access
framing/           fenced tool framing
tools/             OpenAI tool surface and tool-call detection (native actions)
io/                logging, image generation and input, Copilot Studio agent
app/               model-session, credential selection, personal CLI
account-registry.ts  shared: consumed by the protocol, the app module, the chathub
                     code and the proxy's HTTP surface; placement stays open (§Not decided here)
index.ts           unchanged
personal-wiring.ts   composition root: the only module importing both adapters
```

As built, two names differ from the sketch above. The port lives in
`turn-context.ts` rather than `port.ts`, because renaming it would have churned
every importer to change nothing; and `app/credential.ts` is the single selector
the Consequences section requires, sitting beside `model-session` as intended.

The two adapters are `msal` and `personal`. The name `msal` matches the value of
`AuthMode` (`personal-mode.ts:5`) and the `M365_AUTH` environment variable; the
name `personal` is the term the repo already uses for the consumer path. Both are
directory names under `adapters/`, and both are the only two names on that axis.

`app/model-session.ts` is the only place that selects an adapter. `protocol/`
receives a `TurnContext` and never learns which mode produced it.

`index.ts` keeps every export it has today, byte for byte. No package version
changes. `packages/proxy-lib` is untouched in this phase.

Exactly one boundary rule is enforced:

```
adapters/msal must not import adapters/personal, and vice versa
```

`dependency-cruiser` in CI, alongside a test job, is the enforcement mechanism; the
rule set starts from the current graph, violations included, so that movement is
measurable.

## Consequences

- Selecting an adapter happens once. Deleting `auth.ts:686` and `model.ts:75`
  removes the last place where the transport-adjacent code decides auth.
- The shared transport no longer imports account code. `TurnContext.accountId` is
  required and set by the caller; `CopilotSession` never resolves an account.
- Adding a third auth mode means adding one directory under `adapters/` and one
  branch in `app/model-session.ts`. Nothing in `protocol/` changes.
- The protocol cannot be proven unchanged until frame fixtures exist. That work is
  a prerequisite for touching `protocol/`, and it is not in this ADR.

## Status (2026-10-24)

Implemented across #8, #10, #11 and #9, with #2 providing the replay harness the
Consequences section required and #4 pinning the 235-symbol export surface at
177 values and 58 types throughout, so every move proved the public API did not
change.

- `selectCredential` in `app/credential.ts` is the one place the credential is
  chosen; `ModelSession` and the image path both go through it.
- `protocol/session.ts` imports no auth, personal, account or credential module
  and resolves no account: `TurnContext` carries `accountId`, `personal` and the
  record, and the caller supplies it per turn.
- The consumer endpoint (URL, headers, params, payload, probe) lives in
  `personal-chathub/`, so a consumer change no longer edits the enterprise path.
- The boundary rule is enforced by `pnpm lint:boundaries`, reading multi-line and
  dynamic imports, with the graph baselined; `pnpm typecheck` and the unit suite
  run in CI on every push.

Remaining coupling, deliberately left: `personal-chathub/chathub.ts` still imports
`account-registry` for `loginCommandFor`; `personal-chathub/probe.ts` builds its
own socket rather than using the context builder; the personal adapters resolve
their default paths with `resolveAccount()`; and the proxy's request routing
picks the account with `isPersonalMode()` / `resolveAccount()`, which is the host
deciding what a request serves rather than a turn deciding its credential.

- `fenced.ts` (1475 lines), `session.ts` (1230) and `proxy-lib/src/handler.ts`
  (1265) stay large. Splitting them is a separate decision, taken after fixtures
  exist.
- The new modules do not get their own packages. The publish boundary stays at
  `@m365-copilot/core`.
- The barrel remains a 200-symbol firehose. Retiring it for subpath exports is a
  later decision that requires a major version and has no consumer pressure today.

## Rejected alternatives

**Modular monolith by dependency tier.** Rejected for the reasons in Context;
chief among them that its boundary rules required behaviour changes, and that the
tiers had no second implementation to justify them.

**Functional core, imperative shell.** Fits `fenced.ts`, `tools.ts`,
`native-actions.ts` and `priority-access.ts`, which are already pure. It does not
address `session.ts`, which is where the coupling is. It can be applied later
inside `framing/` and `tools/` without contradicting this ADR.

**Clean Architecture layers.** Three layers of entities, use cases and adapters
over 9,000 lines, where no entity concept exists.

**Leaving the current layout and only fixing `TurnContext`.** Cheaper, but leaves
`auth.ts` and `model.ts` both selecting a mode and leaves the chathub code inside
the shared transport, so the next consumer-protocol change touches the
enterprise path again.

## Not decided here

- Whether `accounts` should be a first-class module or part of
  `adapters/personal`. The registry, `resolveAccount` and the CLI/UI routes are
  consumed by both the HTTP surface and the personal adapter.
- Where `account-registry.ts` (913 lines, 33 exports, ~120 lines of test doubles
  in a production file) is split.
- Whether the personal endpoint code deserves its own adapter rather than a
  directory beside `protocol/`.
