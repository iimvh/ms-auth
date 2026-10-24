import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TurnBudgetHold } from "../budget/turn-budget.js";

// ModelSession.run's reconnect: when the first chat() fails it sends the turn
// again on a fresh CopilotSession, and that turn waits on the budget like any
// other (#77). The CopilotSession is faked: each chat() is one upstream turn.
const wire = vi.hoisted(() => ({
  events: [] as string[],
  turnOpts: [] as Array<{ generateImages: boolean } | undefined>,
  contexts: [] as Array<{ plan: { model: string; agentId: string | undefined } }>,
  failures: 0,
}));

vi.mock("../protocol/session.js", () => ({
  CopilotSession: class {
    turnCount = 0;
    constructor(readonly options: Record<string, unknown>) {}
    chat(token: string, _text: string, model: string, _signal: unknown, opts?: { generateImages: boolean }) {
      wire.events.push(`chat:${token}`);
      wire.turnOpts.push(opts);
      const build = this.options.turnContext as (plan: unknown) => unknown;
      build({ requestId: "r", sessionId: "s", conversationId: "c", model, temporaryChat: true, agentId: this.options.agentId, nativeActions: false, images: false });
      wire.contexts.push({ plan: { model, agentId: this.options.agentId as string | undefined } });
      if (wire.failures > 0) {
        wire.failures--;
        return Promise.reject(new Error("WebSocket error: connection failed"));
      }
      return Promise.resolve({ fullText: "ok" });
    }
  },
}));

const { selectCredential } = vi.hoisted(() => ({
  selectCredential: vi.fn(async (accountId?: string | null): Promise<{
    accountId: string | null;
    personal: boolean;
    token: string;
    record: Record<string, unknown> | null;
  }> => ({
    accountId: accountId ?? null,
    personal: false,
    token: "tok1",
    record: null,
  })),
}));
vi.mock("./credential.js", async (importOriginal) => {
  const actual = await importOriginal<Record<string, unknown>>();
  return { ...actual, selectCredential };
});

const PERSONAL_RECORD = {
  accessToken: "tok-personal", chathubPath: "oid@tid", scenario: "S", variants: "v", licenseType: "Free",
  agentHost: "h", product: null, source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
  emailHint: null, probe: null, mintCount: 1, lastMintAt: 0, expiresAt: Date.now() + 3_600_000, capturedAt: 0,
};

const { ModelSession } = await import("./model.js");

function session(accountId?: string | null) {
  let n = 0;
  selectCredential.mockImplementation(async (id?: string | null) => ({
    accountId: id ?? null,
    personal: false,
    token: `tok${++n}`,
    record: null,
  }));
  return new ModelSession({ useAgent: false, accountId });
}

afterEach(() => {
  wire.events = [];
  wire.turnOpts = [];
  wire.contexts = [];
  wire.failures = 0;
  selectCredential.mockReset();
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

  it("is forced OFF for a personal turn", async () => {
    selectCredential.mockImplementation(async (id?: string | null) => ({
      accountId: id ?? null, personal: true, token: "tok-personal", record: PERSONAL_RECORD,
    }));
    await new ModelSession({ useAgent: false }).run("hi", "claude-sonnet", undefined, false);
    expect(wire.turnOpts[0]).toEqual({ generateImages: false, images: [] });
  });

  it("is vetoed by M365_NO_IMAGE_GEN (unchanged)", async () => {
    process.env.M365_NO_IMAGE_GEN = "1";
    await session().run("hi", "claude-sonnet", undefined, false);
    expect(wire.turnOpts[0]).toEqual({ generateImages: false, images: [] });
  });
});

describe("the session's one credential decision", () => {
  afterEach(() => selectCredential.mockReset());

  it("asks the selector once, with the account the session was opened on", async () => {
    await session("work").run("hi", "claude-sonnet", undefined, false);
    expect(selectCredential).toHaveBeenCalledTimes(1);
    expect(selectCredential).toHaveBeenCalledWith("work");
  });

  it("passes the active account (null) when the session pins none", async () => {
    await session().run("hi", "claude-sonnet", undefined, false);
    expect(selectCredential).toHaveBeenCalledWith(null);
  });
});
