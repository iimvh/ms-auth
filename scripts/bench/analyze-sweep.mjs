// Aggregate sweep result JSONs into a strategy × task matrix + per-strategy totals.
// Reads scripts/bench/out/<prefix>-*.json (default prefix "s2"). Each file is one
// (task, strategy) cell. Prints outcome per cell and a ranked leaderboard.
//
//   node scripts/bench/analyze-sweep.mjs [prefix]
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";

const PREFIX = process.argv[2] || "s2";
const OUT = join(process.cwd(), "scripts", "bench", "out");
const files = readdirSync(OUT).filter((f) => f.startsWith(PREFIX + "-") && f.endsWith(".json"));

// label shape: s2-<task>-<strategy>  (task may contain a hyphen: fix-bug, find-needle, edit-config)
const TASKS = ["fix-bug", "find-needle", "edit-config", "fizzbuzz", "count-lines"];
const STRATS = ["baseline", "minimal", "recency", "fewshot", "proof_demand", "persona", "react", "negative", "terse", "reply_tool"];

const OUTCOME_GLYPH = { SOLVED: "✔", GAVE_UP_PROSE: "prose", MAX_TURNS: "maxT", ERROR: "ERR" };

// cell[strategy][task] = latest row for that pair
const cell = {};
for (const f of files) {
  let j;
  try { j = JSON.parse(readFileSync(join(OUT, f), "utf8")); } catch { continue; }
  const label = j.label || "";
  const rest = label.slice(PREFIX.length + 1); // "<task>-<strategy>"
  const task = TASKS.find((t) => rest.startsWith(t + "-"));
  if (!task) continue;
  const strat = rest.slice(task.length + 1);
  const row = (j.rows && j.rows[0]) || {};
  const ts = f.match(/(\d{4}-\d{2}-\d{2}T[\d-]+)\.json$/)?.[1] || "";
  cell[strat] ??= {};
  // keep the most recent capture per cell
  if (!cell[strat][task] || ts > cell[strat][task]._ts) {
    cell[strat][task] = { ...row, _ts: ts, _disengaged: /diseng/i.test(row.error || "") };
  }
}

const usedTasks = TASKS.filter((t) => STRATS.some((s) => cell[s]?.[t]));
const pad = (s, n) => String(s).padEnd(n);

console.log(`\n=== Strategy × Task (prefix ${PREFIX}, ${files.length} files) ===\n`);
console.log(pad("strategy", 13) + usedTasks.map((t) => pad(t, 14)).join("") + "  | solved  tools  diseng");
console.log("-".repeat(13 + usedTasks.length * 14 + 26));
const board = [];
for (const s of STRATS) {
  if (!cell[s]) continue;
  let solved = 0, tools = 0, diseng = 0, cells = 0;
  const cols = usedTasks.map((t) => {
    const r = cell[s][t];
    if (!r) return pad("·", 14);
    cells++;
    if (r.solved) solved++;
    tools += r.toolTurns || 0;
    if (r._disengaged) diseng++;
    const g = r.solved ? "✔SOLVED" : (r._disengaged ? "✗diseng" : (OUTCOME_GLYPH[r.outcome] || r.outcome));
    return pad(`${g}(${r.toolTurns ?? 0}t)`, 14);
  });
  board.push({ s, solved, cells, tools, diseng });
  console.log(pad(s, 13) + cols.join("") + `  | ${solved}/${cells}     ${tools}      ${diseng}`);
}

console.log(`\n=== Leaderboard (by solved, then fewest disengaged) ===`);
board.sort((a, b) => b.solved - a.solved || a.diseng - b.diseng || b.tools - a.tools);
let rank = 1;
for (const b of board) {
  console.log(`  ${rank++}. ${pad(b.s, 13)} solved ${b.solved}/${b.cells}  disengaged ${b.diseng}  toolcalls ${b.tools}`);
}
console.log("");
