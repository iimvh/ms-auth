import { afterEach, describe, expect, it, vi } from "vitest";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
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

vi.mock("ws", async () => ({ default: (await import("./socket-fixture.js")).RecordingWebSocket }));

const { recording, handshake, botText } = await import("./socket-fixture.js");
const { CopilotSession } = await import("./session.js");

const CONTEXT: TurnContext = {
  token: "jwe-token",
  socketUrl: "wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?scenario=S",
  headers: { Origin: "https://m365.cloud.microsoft", "User-Agent": "context-agent" },
  payloadDefaults: { scenario: "OfficeWebIncludedCopilot", licenseType: "Free" },
  capabilities: ["code_interpreter"],
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
    const session = new CopilotSession({ turnContext: CONTEXT });
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
    const session = new CopilotSession({ turnContext: CONTEXT });
    process.env.M365_NO_CODE_INTERPRETER = "1";
    seams.armed = true;

    const pending = session.chat("ignored-token", "hi", "m365-copilot");
    recording.emit("message", handshake());
    await pending;

    expect(seams.touched).toEqual([]);
    expect(recording.sent.join("")).toContain("cwc_code_interpreter");
  });

  it("leaves a session built the old way reading the auth mode", () => {
    process.env.M365_AUTH = "bogus";
    const session = new CopilotSession();
    seams.armed = true;

    expect(() => session.chat("ignored-token", "hi")).toThrow();
    expect(seams.touched).toEqual(["personal-mode.isPersonalMode"]);
  });
});
