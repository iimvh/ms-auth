import { afterEach, describe, expect, it, vi } from "vitest";
import { appendFileSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import {
  awaitTurnBudget,
  budgetConfig,
  parseEvents,
  proxyBudgetStatus,
  readLogEvents,
  recordThrottle,
  recordTurn,
  resetTurnBudget,
  TurnBudgetHold,
  type BudgetConfig,
} from "./turn-budget.js";

const T0 = Date.parse("2026-10-12T08:00:00.000Z");
const iso = (ms: number) => new Date(ms).toISOString();
const turnLine = (ms: number, n = 0) => `[${iso(ms)}] [INFO] [session] Chat turn ${n}: model=claude-sonnet, isFirst=false, text="x"\n`;
const throttleLine = (ms: number) => `[${iso(ms)}] [INFO] [session] Turn result: Throttled (PerUserThrottled)\n`;

function logDir(lines: string[] = []): { dir: string; file: string } {
  const dir = mkdtempSync(join(tmpdir(), "turn-budget-"));
  const file = join(dir, "debug.log");
  writeFileSync(file, lines.join(""));
  return { dir, file };
}

const cfgFor = (dir: string, over: Partial<BudgetConfig> = {}): BudgetConfig => ({
  capacity: 100, refillPerMin: 1.6, reserve: 20, holdMin: 75, windowMin: 360, maxWindowMin: 4320,
  maxWaitMin: 15, enabled: true, dirs: [dir], ...over,
});

function clock(start = T0) {
  const c = { t: start, slept: [] as number[] };
  return {
    c,
    logged: false,
    now: () => c.t,
    sleep: async (ms: number, signal?: AbortSignal) => { signal?.throwIfAborted(); c.slept.push(ms); c.t += ms; },
  };
}

afterEach(() => { vi.useRealTimers(); resetTurnBudget(); });

describe("readLogEvents", () => {
  it("reads only what was appended since the last read", () => {
    const { file } = logDir([turnLine(T0), turnLine(T0 + 1000, 1)]);
    expect(readLogEvents(file)).toHaveLength(2);
    appendFileSync(file, turnLine(T0 + 2000, 2) + throttleLine(T0 + 3000));
    const evs = readLogEvents(file);
    expect(evs.map((e) => e.kind)).toEqual(["turn", "turn", "turn", "throttled"]);
    expect(evs.at(-1)!.t).toBe(T0 + 3000);
  });

  it("holds back a line still being written until its newline arrives", () => {
    const { file } = logDir([turnLine(T0)]);
    const partial = turnLine(T0 + 1000, 1);
    appendFileSync(file, partial.slice(0, 20));
    expect(readLogEvents(file)).toHaveLength(1);
    appendFileSync(file, partial.slice(20));
    expect(readLogEvents(file).map((e) => e.t)).toEqual([T0, T0 + 1000]);
  });

  it("reads a file that shrank from the start again", () => {
    const { file } = logDir([turnLine(T0), turnLine(T0 + 1000, 1), turnLine(T0 + 2000, 2)]);
    expect(readLogEvents(file)).toHaveLength(3);
    writeFileSync(file, turnLine(T0 + 5000));
    expect(readLogEvents(file).map((e) => e.t)).toEqual([T0 + 5000]);
  });

  it("drops events older than keepFrom, and reads again when asked for older ones", () => {
    const { file } = logDir([turnLine(T0), turnLine(T0 + 60_000, 1)]);
    expect(readLogEvents(file, T0 + 30_000).map((e) => e.t)).toEqual([T0 + 60_000]);
    expect(readLogEvents(file, T0).map((e) => e.t)).toEqual([T0, T0 + 60_000]);
  });

  it("returns nothing for a file that's gone", () => {
    const { dir, file } = logDir([turnLine(T0)]);
    expect(readLogEvents(file)).toHaveLength(1);
    rmSync(dir, { recursive: true });
    expect(readLogEvents(file)).toEqual([]);
  });
});

describe("awaitTurnBudget", () => {
  it("reads its limit from M365_BUDGET_MAX_WAIT_MIN", () => {
    expect(budgetConfig({}).maxWaitMin).toBe(15);
    expect(budgetConfig({ M365_BUDGET_MAX_WAIT_MIN: "90" }).maxWaitMin).toBe(90);
  });

  it("is a no-op unless enabled", async () => {
    const { dir } = logDir();
    const res = await awaitTurnBudget({ cfg: cfgFor(dir, { enabled: false }), logged: false, sleep: () => { throw new Error("slept"); } });
    expect(res.waitedMs).toBe(0);
  });

  it("lets a turn through at once with room above the reserve", async () => {
    const { dir } = logDir();
    const k = clock();
    expect((await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep })).waitedMs).toBe(0);
    expect(k.c.slept).toEqual([]);
  });

  it("waits on a drained bucket until it holds one turn plus the reserve, re-reading every 15 s", async () => {
    const { dir } = logDir(Array.from({ length: 85 }, (_, i) => turnLine(T0 - 85_000 + i * 1000, i)));
    const k = clock();
    const res = await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep });
    expect(k.c.slept.every((ms) => ms <= 15_000)).toBe(true);
    expect(res.waitedMs).toBeGreaterThan(2 * 60_000);
    expect(res.waitedMs).toBeLessThan(3 * 60_000);
  });

  it("counts a turn it let through before it starts, so the next request waits for it", async () => {
    const { dir } = logDir(Array.from({ length: 79 }, (_, i) => turnLine(T0 - 79 + i, i)));
    const k = clock();
    const cfg = cfgFor(dir);
    await awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep });
    expect(k.c.slept).toEqual([]);
    await awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep });
    expect(k.c.slept.length).toBeGreaterThan(0);
    const before = proxyBudgetStatus(k.now(), cfg, false).turns;
    recordTurn(k.now());
    expect(proxyBudgetStatus(k.now(), cfg, false).turns).toBe(before);
  });

  it("decides one request at a time", async () => {
    const { dir } = logDir(Array.from({ length: 79 }, (_, i) => turnLine(T0 - 79 + i, i)));
    const k = clock();
    const cfg = cfgFor(dir);
    const [a, b] = await Promise.all([
      awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep }),
      awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep }),
    ]);
    expect(a.waitedMs).toBe(0);
    expect(b.waitedMs).toBeGreaterThan(0);
  });

  it("counts its own turns when it writes no log, and leaves them to the log when it does", async () => {
    const at = Date.parse("2020-01-01T00:00:00.000Z");
    const { dir } = logDir();
    const cfg = cfgFor(dir);
    for (let i = 0; i < 85; i++) recordTurn(at - 85 + i);
    const st = proxyBudgetStatus(at, cfg, false);
    expect(st.turns).toBe(85);
    expect(st.level).toBeLessThan(21);
    expect(proxyBudgetStatus(at, cfg, true).turns).toBe(0);
  });

  it("answers a throttle hold with TurnBudgetHold instead of waiting it out", async () => {
    const { dir } = logDir([turnLine(T0 - 10 * 60_000), throttleLine(T0 - 10 * 60_000)]);
    const k = clock();
    const err = await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep }).catch((e) => e);
    expect(err).toBeInstanceOf(TurnBudgetHold);
    expect(err.reason).toBe("throttle hold");
    expect(Math.round(err.waitMs / 60_000)).toBe(65);
    expect(err.lastThrottle).toBe(T0 - 10 * 60_000);
    expect(k.c.slept).toEqual([]);
  });

  it("sees a throttle this process recorded", async () => {
    const { dir } = logDir();
    recordTurn(T0 - 1000);
    recordThrottle(T0 - 1000);
    const k = clock();
    const err = await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep }).catch((e) => e);
    expect(err).toBeInstanceOf(TurnBudgetHold);
    expect(err.reason).toBe("throttle hold");
  });

  it("waits out the end of a hold that fits in the request's limit", async () => {
    const { dir } = logDir([throttleLine(T0 - 70 * 60_000)]);
    const k = clock();
    const res = await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep });
    expect(Math.round(res.waitedMs / 60_000)).toBe(5);
  });

  it("refuses a refill longer than the request may wait, and the limit is the smaller of the two", async () => {
    const { dir } = logDir(Array.from({ length: 100 }, (_, i) => turnLine(T0 - 100 + i, i)));
    const k = clock();
    const err = await awaitTurnBudget({ cfg: cfgFor(dir), maxWaitMs: 2 * 60_000, logged: k.logged, now: k.now, sleep: k.sleep }).catch((e) => e);
    expect(err).toBeInstanceOf(TurnBudgetHold);
    expect(err.reason).toBe("refill");
    const err2 = await awaitTurnBudget({ cfg: cfgFor(dir, { maxWaitMin: 1 }), maxWaitMs: 60 * 60_000, logged: k.logged, now: k.now, sleep: k.sleep }).catch((e) => e);
    expect(err2).toBeInstanceOf(TurnBudgetHold);
    const res = await awaitTurnBudget({ cfg: cfgFor(dir), logged: k.logged, now: k.now, sleep: k.sleep });
    expect(res.waitedMs).toBeGreaterThan(12 * 60_000);
  });

  it("stops waiting when the client goes away, and lets the next request decide", async () => {
    const { dir } = logDir(Array.from({ length: 90 }, (_, i) => turnLine(T0 - 90 + i, i)));
    const k = clock();
    const ac = new AbortController();
    const cfg = cfgFor(dir);
    const waiting = awaitTurnBudget({ cfg, logged: false, now: k.now, signal: ac.signal, sleep: (ms, signal) => new Promise((_, rej) => signal?.addEventListener("abort", () => rej(signal.reason))) });
    ac.abort(new Error("client gone"));
    await expect(waiting).rejects.toThrow("client gone");
    const res = await awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep });
    expect(res.waitedMs).toBeGreaterThan(0);
  });
});

describe("awaitTurnBudget's line (#77)", () => {
  const drained = () => logDir(Array.from({ length: 85 }, (_, i) => turnLine(T0 - 85_000 + i * 1000, i)));

  function heldFront(k: ReturnType<typeof clock>, cfg: BudgetConfig) {
    let release!: () => void;
    const gate = new Promise<void>((res) => { release = () => { k.c.t += 10 * 60_000; res(); }; });
    const done = awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: () => gate });
    return { done, release };
  }

  const settled = async (p: Promise<unknown>) => {
    let out = false;
    void p.then(() => { out = true; }, () => { out = true; });
    for (let i = 0; i < 20; i++) await Promise.resolve();
    return out;
  };

  it("counts the time behind another request toward its own wait and limit", async () => {
    const { dir } = drained();
    const cfg = cfgFor(dir);
    const later = (k: ReturnType<typeof clock>) => async (ms: number, signal?: AbortSignal) => { await null; await k.sleep(ms, signal); };
    const k = clock();
    const [a, b] = await Promise.all([
      awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: later(k) }),
      awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: later(k) }),
    ]);
    expect(Math.round(a.waitedMs / 1000)).toBe(140);
    expect(Math.round((b.waitedMs - a.waitedMs) / 100) / 10).toBe(37.5);
    resetTurnBudget();
    const k2 = clock();
    const [, err] = await Promise.all([
      awaitTurnBudget({ cfg, logged: k2.logged, now: k2.now, sleep: later(k2) }),
      awaitTurnBudget({ cfg, maxWaitMs: 2.75 * 60_000, logged: k2.logged, now: k2.now, sleep: later(k2) }).catch((e) => e),
    ]);
    expect(err).toBeInstanceOf(TurnBudgetHold);
    expect(err.reason).toBe("refill");
  });

  it("leaves the line at its own limit while the request ahead keeps waiting", async () => {
    vi.useFakeTimers({ toFake: ["setTimeout", "clearTimeout"] });
    const { dir } = drained();
    const cfg = cfgFor(dir);
    const k = clock();
    const front = heldFront(k, cfg);
    const queued = awaitTurnBudget({ cfg, maxWaitMs: 2 * 60_000, logged: k.logged, now: k.now, sleep: k.sleep }).catch((e) => e);
    await vi.advanceTimersByTimeAsync(2 * 60_000 - 1);
    expect(await settled(queued)).toBe(false);
    await vi.advanceTimersByTimeAsync(1);
    const err = await queued;
    expect(err).toBeInstanceOf(TurnBudgetHold);
    expect(err.reason).toBe("refill");
    expect(err.waitMs).toBeGreaterThan(2.5 * 60_000);
    front.release();
    expect((await front.done).waitedMs).toBe(10 * 60_000);
    expect((await awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep })).waitedMs).toBe(0);
  });

  it("leaves the line at once when its client goes away, and the next one still waits for the front", async () => {
    const { dir } = drained();
    const cfg = cfgFor(dir);
    const k = clock();
    const front = heldFront(k, cfg);
    const ac = new AbortController();
    const gone = awaitTurnBudget({ cfg, signal: ac.signal, logged: k.logged, now: k.now, sleep: k.sleep });
    const next = awaitTurnBudget({ cfg, logged: k.logged, now: k.now, sleep: k.sleep });
    ac.abort(new Error("client gone"));
    await expect(gone).rejects.toThrow("client gone");
    expect(await settled(next)).toBe(false);
    expect(k.c.slept).toEqual([]);
    front.release();
    expect((await front.done).waitedMs).toBe(10 * 60_000);
    expect((await next).waitedMs).toBe(10 * 60_000);
  });
});

describe("account-tagged turn events", () => {
  it("counts one account's turns only against that account", () => {
    const log = [
      "[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a account=work",
      "[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a account=work",
      "[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a",
    ].join("\n");
    expect(parseEvents(log, "work")).toHaveLength(2);
    expect(parseEvents(log)).toHaveLength(1);
  });

  it("an untagged log (written before this change) still counts", () => {
    expect(parseEvents("[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a")).toHaveLength(1);
  });

  it("tags a throttle line the same way", () => {
    const log = "[2026-10-23T05:00:00.000Z] [INFO] [session] Turn result: Throttled (PerUserThrottled) account=work";
    expect(parseEvents(log, "work")).toEqual([expect.objectContaining({ kind: "throttled" })]);
    expect(parseEvents(log, "other")).toEqual([]);
  });

  it("reads the tag as the last token, not one inside the prompt text", () => {
    const spoof = JSON.stringify("please note account=work in the ticket");
    const line = `[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a, isFirst=true, text=${spoof} account=default`;
    expect(parseEvents(line, "default")).toEqual([expect.objectContaining({ kind: "turn", account: "default" })]);
    expect(parseEvents(line, "work")).toEqual([]);
    expect(parseEvents(line)).toEqual([]);
  });

  it("an untagged line whose text mentions an account is the builtin account's", () => {
    const line = '[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a, isFirst=true, text="account=work"';
    expect(parseEvents(line)).toEqual([expect.objectContaining({ kind: "turn", account: null })]);
    expect(parseEvents(line, "work")).toEqual([]);
  });

  it("a throttled line with free text and a trailing tag still counts", () => {
    const line = '[2026-10-23T05:00:00.000Z] [INFO] [session] Turn result: Throttled (PerUserThrottled) message="account=other" account=work';
    expect(parseEvents(line, "work")).toEqual([expect.objectContaining({ kind: "throttled" })]);
    expect(parseEvents(line, "other")).toEqual([]);
  });

  it("hashes two account ids of equal length as different buckets", () => {
    const a = "[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a account=" + "a".repeat(26);
    const b = "[2026-10-23T05:00:00.000Z] [INFO] [session] Chat turn 0: model=a account=" + "b".repeat(26);
    expect(parseEvents(a, "a".repeat(26))).toHaveLength(1);
    expect(parseEvents(b, "a".repeat(26))).toHaveLength(0);
  });
});
