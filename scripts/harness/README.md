# Tool-call test harness

Answers one question empirically: **which models + system-prompt sizes + toolset sizes
actually do tool calls correctly through the proxy** — and where a big system prompt
kills tool-calling (a failure mode seen repeatedly with real coding-agent prompts).

It drives the proxy as a real OpenAI tool loop over a Docker sandbox, runs each task's
objective verifier, and records three metrics per cell:

- **compliance** — did the model emit a tool call on turn 1 (vs prose / give-up)
- **solve** — did the task's verifier pass (the code actually works / right answer)
- **disengage** — did M365's safety filter refuse (the proxy surfaces a 502 `disengaged`)

## Dimensions

| axis | values | where |
|---|---|---|
| model | any proxy model id (`m365-copilot`, `gpt-5.5-think-deeper`, `claude`, …) | `MODELS` |
| system-prompt size | `none` `small` `medium` `large` `huge` (60 → 26 000 chars) | `prompts/sys_*.txt` |
| toolset size | `lean` (1) `standard` (4) `large` (12, opencode-like) | `--tool-preset` |
| task | verifiable bench tasks (`fix-bug`, `fizzbuzz`, `find-needle`, `edit-config`, …) | `TASKS` |

Prompt fixtures are generated (reproducible sizes): `node scripts/harness/prompts/gen.mjs`.

## Run

```bash
# small default (1 model × 5 prompt sizes × standard toolset × fix-bug)
node scripts/harness/matrix.mjs

# hunt the failure modes: 2 models × extremes × the big toolset
MODELS="m365-copilot,gpt-5.5-think-deeper" PROMPTS="none,large,huge" \
  PRESETS="lean,standard,large" TASKS="fix-bug,edit-config" node scripts/harness/matrix.mjs

# read the grid + prompt-size "death curve"
node scripts/harness/analyze-matrix.mjs
```

`matrix.mjs` starts an in-process proxy (`serve.mjs`, no Nitro build needed), runs the
grid via `run-cell.mjs`, and cleans up. Each cell writes a JSON to `out/`.

## Cost / quota

Each cell = `#tasks × REPEAT` multi-turn tool loops → several M365 messages. Total ≈
`#models × #prompts × #presets` cells. The default is deliberately small. The driver
paces between cells; a fresh nonce per task means each task is its own M365 conversation
(so per-conversation quota isn't the limit — the thread-rate throttle is, if you sweep
very wide). Scale deliberately; run on a rested account for a big sweep.

## Prerequisites

- Docker (sandbox: `--network none`, task dir mounted, runs as host uid)
- A logged-in M365 proxy (auth via the repo's cached login; `serve.mjs` warms it)

## First findings (2026-07-13, n=1/cell — illustrative, not conclusive)

- `gpt-5.5-think-deeper`: **100% compliance + 100% solve** on `fix-bug` across `none`↔`huge`
  prompt (311↔26 201 chars) and `standard`↔`large` toolset. No disengage. Robust.
- `m365-copilot` (default magic/GPT-tone + agent): 100% compliance but **0% solve** on the
  same — it calls tools but gets the answer wrong. No disengage in these cells.
- Takeaway so far: the reasoning model is the reliable one; big prompts did **not** break
  compliance here. Widen the sweep (more tasks/models, `REPEAT>1`) to locate the combos
  where they do.
