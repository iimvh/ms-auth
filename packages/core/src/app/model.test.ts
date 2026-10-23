import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TurnBudgetHold } from "../budget/turn-budget.js";

// ModelSession.run's reconnect: when the first chat() fails it sends the turn
// again on a fresh CopilotSession, and that turn waits on the budget like any
// other (#77). The CopilotSession is faked: each chat() is one upstream turn.
const wire = vi.hoisted(() => ({ events: [] as string[], turnOpts: [] as Array<{ generateImages: boolean } | undefined>, failures: 0 }));

vi.mock("../protocol/session.js", () => ({
  CopilotSession: class {
    turnCount = 0;
    chat(token: string, _text: string, _model: string, _signal: unknown, opts?: { generateImages: boolean }) {
      wire.events.push(`chat:${token}`);
      wire.turnOpts.push(opts);
      if (wire.failures > 0) {
        wire.failures--;
        return Promise.reject(new Error("WebSocket error: connection failed"));
      }
      return Promise.resolve({ fullText: "ok" });
    }
  },
}));

const { ModelSession } = await import("./model.js");

function session() {
  let n = 0;
  return new ModelSession({ useAgent: false, getToken: async () => `tok${++n}` });
}

afterEach(() => {
  wire.events = [];
  wire.turnOpts = [];
  wire.failures = 0;
});

describe("ModelSession.run's reconnect", () => {
  it("waits on the turn budget before it sends the turn again, with a fresh token after a wait", async () => {
    wire.failures = 1;
    const waitForBudget = async () => { wire.events.push("wait"); return { waitedMs: 30_000 }; };
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(wire.events).toEqual(["chat:tok1", "wait", "chat:tok2"]);
  });

  it("keeps the token when the budget let it through at once", async () => {
    wire.failures = 1;
    const waitForBudget = async () => { wire.events.push("wait"); return { waitedMs: 0 }; };
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(wire.events).toEqual(["chat:tok1", "wait", "chat:tok1"]);
  });

  it("doesn't send it when the budget refuses the turn", async () => {
    wire.failures = 1;
    const hold = new TurnBudgetHold("refill", 5 * 60_000, null);
    const waitForBudget = async () => { throw hold; };
    await expect(session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget)).rejects.toBe(hold);
    expect(wire.events).toEqual(["chat:tok1"]);
  });

  it("doesn't wait again when the first turn goes through", async () => {
    const waitForBudget = vi.fn(async () => ({ waitedMs: 0 }));
    await session().run("hi", "claude-sonnet", undefined, false, [], waitForBudget);
    expect(waitForBudget).not.toHaveBeenCalled();
    expect(wire.events).toEqual(["chat:tok1"]);
  });

  it("waits with awaitTurnBudget by default (a no-op unless M365_AVOID_THROTTLING is set)", async () => {
    const saved = process.env.M365_AVOID_THROTTLING;
    delete process.env.M365_AVOID_THROTTLING;
    try {
      wire.failures = 1;
      await session().run("hi", "claude-sonnet", undefined, false);
      expect(wire.events).toEqual(["chat:tok1", "chat:tok1"]);
    } finally {
      if (saved !== undefined) process.env.M365_AVOID_THROTTLING = saved;
    }
  });
});

describe("generateImages defaults", () => {
  beforeEach(() => {
    delete process.env.M365_AUTH;
    delete process.env.M365_NO_IMAGE_GEN;
  });
  afterEach(() => {
    delete process.env.M365_AUTH;
    delete process.env.M365_NO_IMAGE_GEN;
  });

  it("stays ON for an agent-less msal turn (today's behaviour, byte for byte)", async () => {
    await session().run("hi", "claude-sonnet", undefined, false);
    expect(wire.turnOpts[0]).toEqual({ generateImages: true, images: [] });
  });

  it("is forced OFF in personal mode", async () => {
    process.env.M365_AUTH = "personal";
    await session().run("hi", "claude-sonnet", undefined, false);
    expect(wire.turnOpts[0]).toEqual({ generateImages: false, images: [] });
  });

  it("is vetoed by M365_NO_IMAGE_GEN (unchanged)", async () => {
    process.env.M365_NO_IMAGE_GEN = "1";
    await session().run("hi", "claude-sonnet", undefined, false);
    expect(wire.turnOpts[0]).toEqual({ generateImages: false, images: [] });
  });
});
