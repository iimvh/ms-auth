import { afterEach, describe, expect, it, vi } from "vitest";
import type { TurnContext } from "./turn-context.js";

const seams = vi.hoisted(() => ({
  armed: false,
  touched: [] as string[],
  wrap(prefix: string, actual: Record<string, unknown>): Record<string, unknown> {
    return new Proxy(actual, {
      get: (target, key, receiver) => {
        const value = Reflect.get(target, key, receiver);
        if (!seams.armed || typeof value !== "function") return value;
        return (..._args: unknown[]) => {
          seams.touched.push(`${prefix}.${String(key)}`);
          throw new Error(`${prefix}.${String(key)} was reached`);
        };
      },
    });
  },
}));

vi.mock("./adapters/personal/personal-harvest-cdp.js", () => ({ attachChatPage: () => Promise.reject(new Error("no browser")) }));
vi.mock("./adapters/personal/personal-auth.js", () => ({
  PersonalAuthError: class PersonalAuthError extends Error {},
  personalLoginCommand: () => "pnpm m365-personal login",
}));
vi.mock("./account-registry.js", async (importOriginal) =>
  seams.wrap("account-registry", await importOriginal<Record<string, unknown>>()));
vi.mock("./adapters/personal/personal-mode.js", async (importOriginal) =>
  seams.wrap("personal-mode", await importOriginal<Record<string, unknown>>()));
vi.mock("./adapters/msal/auth.js", async (importOriginal) =>
  seams.wrap("msal-auth", await importOriginal<Record<string, unknown>>()));
vi.mock("./personal-chathub/chathub.js", async (importOriginal) =>
  seams.wrap("chathub", await importOriginal<Record<string, unknown>>()));

vi.mock("ws", async () => ({ default: (await import("./protocol/socket-fixture.js")).RecordingWebSocket }));

const { recording, handshake, botText } = await import("./protocol/socket-fixture.js");
const { CopilotSession } = await import("./protocol/session.js");

const CONTEXT: TurnContext = {
  token: "jwe-token",
  accountId: "work",
  personal: true,
  socketUrl: "wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?scenario=S",
  headers: { Origin: "https://m365.cloud.microsoft", "User-Agent": "context-agent" },
  payloadDefaults: { scenario: "OfficeWebIncludedCopilot", licenseType: "Free" },
  capabilities: ["code_interpreter"],
  record: null,
};
const ANSWER = "hello from the context";

afterEach(() => {
  seams.armed = false;
  seams.touched.length = 0;
  delete process.env.M365_AUTH;
  delete process.env.M365_NO_CODE_INTERPRETER;
});

describe("a session with a turn context (#8)", () => {
  it("performs a turn from the context without the auth mode or an account", async () => {
    const session = new CopilotSession({ turnContext: () => CONTEXT, accountId: "work" });
    process.env.M365_AUTH = "bogus";
    seams.armed = true;

    const pending = session.chat("ignored-token", "hi", "claude-sonnet");
    recording.emit("message", handshake());
    recording.emit("message", botText(ANSWER));
    const stream = await pending;
    recording.emit("close", Buffer.from(""));
    let text = "";
    for await (const delta of stream) text += delta;

    expect(text).toBe(ANSWER);
    expect(stream.fullText).toBe(ANSWER);
    expect(session.turnCount).toBe(1);
    expect(seams.touched).toEqual([]);
    expect(recording.url).toBe(CONTEXT.socketUrl);
    expect(recording.options).toEqual({ headers: CONTEXT.headers });
    const payload = recording.sent.join("");
    expect(payload).toContain('"scenario":"OfficeWebIncludedCopilot"');
    expect(payload).toContain('"tone":"Claude_Sonnet"');
    expect(payload).toContain("cwc_code_interpreter");
  });

  it("takes its capabilities from the context, not from the operator's environment", async () => {
    const session = new CopilotSession({ turnContext: () => CONTEXT });
    process.env.M365_NO_CODE_INTERPRETER = "1";
    seams.armed = true;

    const pending = session.chat("ignored-token", "hi", "m365-copilot");
    recording.emit("message", handshake());
    await pending;

    expect(seams.touched).toEqual([]);
    expect(recording.sent.join("")).toContain("cwc_code_interpreter");
  });

  it("builds the context afresh for every turn, so a context builder is asked per turn", async () => {
    const plans: Array<{ sessionId: string; model: string; agentId: string | undefined }> = [];
    const session = new CopilotSession({
      sessionId: "sid-fixed",
      conversationId: "cid-fixed",
      agentId: "agent-1",
      turnContext: (plan) => {
        plans.push({ sessionId: plan.sessionId, model: plan.model, agentId: plan.agentId });
        return CONTEXT;
      },
    });

    const first = session.chat("ignored-token", "one", "claude-sonnet");
    recording.emit("message", handshake());
    await first;
    const second = session.chat("ignored-token", "two", "m365-copilot");
    recording.emit("message", handshake());
    await second;
    expect(plans).toEqual([
      { sessionId: "sid-fixed", model: "claude-sonnet", agentId: "agent-1" },
      { sessionId: "sid-fixed", model: "m365-copilot", agentId: "agent-1" },
    ]);
  });

  it("refuses to be built without a context: there is no mode-reading fallback left", () => {
    process.env.M365_AUTH = "bogus";
    seams.armed = true;

    // @ts-expect-error a session without a context is a compile-time error
    expect(() => new CopilotSession()).toThrow();
    expect(seams.touched).toEqual([]);
  });
});
