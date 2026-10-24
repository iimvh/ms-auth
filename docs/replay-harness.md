# Frame replay harness (#2)

Replay a recorded M365 turn through the transport and stream code without
contacting M365. The harness is the fixture format, its loader, and a test that
drives the same path a live turn takes.

## Format: the proxy's own dump, not a second one

The reader consumes exactly what `M365_DUMP_FRAMES=1` writes. `dumpFrame()` in
`packages/core/src/session.ts` appends one NDJSON line per frame to
`<FRAME_DIR>/<requestId>.ndjson`:

```json
{"t":1792790402530,"dir":"recv","frame":{"type":1,"target":"update","arguments":[{"writeAtCursor":" through","streamingMode":"Delta"}]}}
```

| field | meaning |
|---|---|
| `t` | epoch ms; the loader keeps the offset from the first line so a replay can pace frames as they arrived |
| `dir` | `recv` (delivered to the turn) or `send` (written by the turn) |
| `frame` | the parsed SignalR frame, verbatim |
| `synthetic` | optional note; **every** line of a hand-written fixture carries it, a recorded dump carries none |

`#3` delivers real recordings in this shape for free; nothing changes here when
they land.

## Curation

A live dump records everything on the wire, and some of it cannot be replayed
faithfully: a `M365_DUMP_FRAMES` send line keeps only the frame, not the record
separator the socket wrote, and a handshake response is never captured. A
replayable fixture is therefore a **curated** dump: received frames kept as the
dump wrote them, the turn's own sends dropped, and a leading `send` line holding
the SignalR handshake so the turn opens the way it did live. The synthetic
fixture follows that rule (`packages/core/fixtures/replay/`).

## Where fixtures live, and why

`packages/core/fixtures/replay/` — inside the package that owns the reader, the
socket fixture and the replay test, not `scripts/`.

- The reader must stay importable from the core test tree; `scripts/` sits
  outside every package's `tsconfig` include and is not built by tsdown, so a
  loader there could not be unit-tested with the code it parses.
- `scripts/bench/turn-budget.test.mjs` is the precedent for a *re-export* test
  next to a script's consumer, but that script is exercised by the bench. A
  fixture is data for a core test, and core's own file layout seats it beside
  `socket-fixture.ts`.
- Keeping it under version control means every real recording `#3` adds is a
  reviewable diff of protocol behaviour, next to the code that parses it.

## Loader API (`packages/core/src/frame-replay.ts`)

| export | purpose |
|---|---|
| `FrameDumpLine` | zod schema for one dump line; validates a recording before replay |
| `loadFrameFixture(text, file?)` | parse a dump into a `ReplayFixture` |
| `readFrameFixture(path)` | the same, from a file |
| `FrameFixtureError` | a line that is not a dump line, an empty dump, or a mixed synthetic/recorded file |

`ReplayFixture` carries `lines` (every parsed line), `received` (one SignalR
message buffer per received frame, in order, each with its `at` offset and the
dump `line` it came from), `synthetic` (the note all lines share, or `null`),
and `file`.

`socket-fixture.ts` gains `replayFeed(frames)`, which emits each loaded frame to
the `RecordingWebSocket` as one `message` event — the path `handshake()` +
`botText()` only approximated.

## What the replay test protects

`packages/core/src/frame-replay.test.ts` drives `CopilotSession.chat()` through
a `TurnContext` (the `chatWithTurnContext` path, the same one a live turn takes)
over the recording socket. The synthetic fixture holds a complete turn: the
handshake, an empty `Progress` frame, a first bot message streamed as a
head snapshot plus deltas plus a final snapshot, a **second** bot message whose
head arrives **only** as a `cursor` snapshot then deltas, a message update, the
`type:2` completion and the `type:3`/`type:7` teardown.

Perturbing any one of these turns the replay test red:

| change to `session.ts` | failing assertion |
|---|---|
| `TurnTextComposer.snapshot` folds every message into one slot | `expected '```bash\ncat notes/secret.txt\n```\n…' to be 'Let me look through the files in note…'` |
| the cursor id is ignored (`composer.cursor` never called) | `expected 'Let me look through the files in notes/.bash\n\n```bash…' to be 'Let me look through the files in notes/.\n\n```bash…'` |
| `onAbort` no longer sends the Stop frame | `expected '{"arguments":[{"scenario":"OfficeWebI…' to contain '"target":"stop"'` |

## Recorded M365 turns (#3)

Three real captures recorded on 2026-10-24 through a personal (MSA) account in
`OfficeWebFreeConsumerCopilot`, replayed by
`packages/core/src/protocol/frame-replay-recorded.test.ts`:

| fixture | what M365 sent |
|---|---|
| `recorded-plain-answer.ndjson` | a one-sentence answer streamed as a cursor head snapshot ("Paris") plus two deltas, a separate throttling update, then `ReferencesListComplete`, the type:2 completion with the user echo, and type:3 |
| `recorded-throttling-update.ndjson` | a refusal-to-edit answer with the sandbox frames (`contentType:"Code"`, `hiddenText` bash runs), the throttling update, and a type:7 `allowReconnect` after type:3 |
| `recorded-disengaged.ndjson` | the filter firing: an "Apology" message streamed as text, a `messageType:"Disengaged"` update (`contentOrigin:"DeepLeo"`, `hiddenText:"> Conversation disengaged"`), and a second Disengaged marker (`contentOrigin:"JailBreakClassifier"`, `hiddenText:"Disengaged by JailBreak Classifier"`) in the type:2 completion |

The recordings pin what the synthetic fixture approximates:

- the real server says **Disengaged comes in an update frame** (before type:2)
  *and* repeats the marker in the completion item, so `stream.messageType`
  ends as `"Disengaged"` (last write) — the open question #3's earlier review
  could not settle from the repo;
- the apology **is** streamed as content (`contentOrigin:"Apology"` with no
  `messageType`, so the `!m.messageType` snapshot gate passes it);
- throttling arrives as its own update argument with all three numbers, and a
  second, richer copy (with `metering`) rides the type:2 item.

The captures are real dumps, not hand-written: only the `locationInfo` blobs
M365 attaches to the user echo were blanked (the repo is public), everything
else is byte-for-byte what the proxy wrote.

## Running it

```sh
pnpm build && npx vitest run packages/core/src/frame-replay.test.ts
```

No M365 request is made: `ws` is mocked with `RecordingWebSocket` and no real
credential or account is touched, so no quota is spent.

```sh
npx vitest run packages/core/src/protocol/frame-replay-recorded.test.ts
```

replays the three recorded captures through the same path (also offline).
