import { describe, expect, it, vi } from "vitest";
import { fileURLToPath } from "node:url";
import type { TurnContext } from "../turn-context.js";

vi.mock("../adapters/personal/personal-harvest-cdp.js", () => ({ attachChatPage: () => Promise.reject(new Error("no browser")) }));
vi.mock("../adapters/personal/personal-auth.js", () => ({
  PersonalAuthError: class PersonalAuthError extends Error {},
  personalLoginCommand: () => "pnpm m365-personal login",
}));

vi.mock("ws", async () => ({ default: (await import("./socket-fixture.js")).RecordingWebSocket }));

const { recording, replayFeed } = await import("./socket-fixture.js");
const { readFrameFixture } = await import("./frame-replay.js");
const { CopilotSession } = await import("./session.js");

const FIXTURE = fileURLToPath(new URL("../../fixtures/replay/synthetic-complete-turn.ndjson", import.meta.url));

const CONTEXT: TurnContext = {
  token: "jwe-token",
  socketUrl: "wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?scenario=S",
  headers: { Origin: "https://m365.cloud.microsoft", "User-Agent": "replay-agent" },
  payloadDefaults: { scenario: "OfficeWebIncludedCopilot", licenseType: "Free" },
  capabilities: ["code_interpreter"],
};

const NARRATION = "Let me look through the files in notes/.";
const FENCE_HEAD = "```";
const FENCE = FENCE_HEAD + "bash\ncat notes/secret.txt\n```";
const INVENTED = "\n\n<tool_response>\nthe secret is xyzzy\n</tool_response>";
const SINGLE_MESSAGE_FOLD = `${NARRATION}${FENCE}${INVENTED}\n\nNow I run it:`;
const ANSWER = `${NARRATION}\n\n${FENCE}${INVENTED}\n\nNow I run it:`;

async function replayTurn() {
  const fixture = readFrameFixture(FIXTURE);
  const session = new CopilotSession({ turnContext: CONTEXT });
  const pending = session.chat("ignored-token", "grep the notes directory for the secret and run it", "claude-sonnet");
  replayFeed(fixture.received)();
  const stream = await pending;
  recording.emit("close", Buffer.from(""));
  const deltas: string[] = [];
  for await (const delta of stream) deltas.push(delta);
  return { fixture, session, stream, deltas, streamed: deltas.join(""), sent: recording.sent.join("") };
}

describe("replaying a complete turn from a frame dump (#2)", () => {
  it("loads a dump written in the proxy's own format, and it says it is synthetic", () => {
    const fixture = readFrameFixture(FIXTURE);
    expect(fixture.synthetic).toBe("synthetic, not a recording");
    expect(fixture.lines).toHaveLength(19);
    expect(fixture.lines.filter((line) => line.dir === "send")).toHaveLength(2);
    expect(fixture.received).toHaveLength(17);
    expect(fixture.received.map((frame) => frame.line)).toEqual([2, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]);
    expect(fixture.received[0]!.at).toBe(40);
    expect(fixture.received.at(-1)!.at).toBe(5045);
    const times = fixture.received.map((frame) => frame.at);
    expect(times).toEqual([...times].sort((a, b) => a - b));
  });

  it("carries a second message whose head is a cursor snapshot, not a delta", () => {
    const fixture = readFrameFixture(FIXTURE);
    const update = (line: number) => JSON.parse(fixture.received.find((frame) => frame.line === line)!.data.toString().slice(0, -1));
    const head = update(10);
    expect(head.arguments[0].cursor.j).toBe("$['9d3b0c62-7777-4f58-8b21-2e5a6d7c8f90'].adaptiveCards[0].body[0].text");
    expect(head.arguments[0].messages[0]).toEqual({
      text: "```",
      author: "bot",
      messageId: "9d3b0c62-7777-4f58-8b21-2e5a6d7c8f90",
      requestId: "fixture-request-1",
      contentOrigin: "DeepLeo",
    });
    const deltas = [11, 12, 13, 15].map((line) => update(line).arguments[0].writeAtCursor as string);
    expect(deltas.join("")).toBe("bash\ncat notes/secret.txt\n```" + INVENTED + "\n\nNow I run it:");
    expect(FENCE_HEAD + deltas.join("")).toBe(FENCE + INVENTED + "\n\nNow I run it:");
  });

  it("rejects a dump that is neither wholly synthetic nor wholly recorded", async () => {
    const { loadFrameFixture } = await import("./frame-replay.js");
    const lines = readFrameFixture(FIXTURE).lines;
    const mixed = lines.map((line, index) => ({ ...line, synthetic: undefined })).slice(0, 1).concat(lines.slice(1));
    const text = mixed.map((line) => JSON.stringify(line)).join("\n");
    expect(() => loadFrameFixture(text, "mixed.ndjson")).toThrow(/synthetic/);
  });

  it("composes both messages, keeping the second one's cursor-snapshot head", async () => {
    const { stream, streamed } = await replayTurn();
    expect(stream.fullText).toBe(ANSWER);
    expect(streamed).toBe(ANSWER);
    expect(stream.fullText).not.toBe(SINGLE_MESSAGE_FOLD);
    expect(stream.fullText).toContain("notes/.\n\n```bash");
    expect(stream.fullText.startsWith(FENCE)).toBe(false);
  });

  it("drives the turn over the socket the context names, sending the chat invocation", async () => {
    const { sent, session } = await replayTurn();
    expect(recording.url).toBe(CONTEXT.socketUrl);
    const messages = sent.split("\x1E").filter(Boolean);
    expect(messages.filter((message) => message.includes('"target":"chat"'))).toHaveLength(1);
    expect(messages.find((message) => message.includes('"target":"chat"'))).toContain('"invocationId":"0"');
    expect(sent).toContain('"scenario":"OfficeWebIncludedCopilot"');
    expect(sent).toContain('"tone":"Claude_Sonnet"');
    expect(session.turnCount).toBe(1);
  });

  it("takes the answer from the streamed messages, and the turn's metadata stays unset on this path", async () => {
    const { stream } = await replayTurn();
    expect(stream.hasContent).toBe(true);
    expect(stream.messageId).toBeNull();
    expect(stream.contentOrigin).toBeNull();
    expect(stream.messageType).toBeNull();
    expect(stream.turnState).toBeNull();
    expect(stream.result).toBeNull();
    expect(stream.throttle).toBeNull();
    expect(stream.scores).toBeNull();
  });

  it("ends the turn on the completion frame, before the dump's stop apology", async () => {
    const { stream } = await replayTurn();
    expect(stream.fullText).not.toContain("You have stopped");
    expect(stream.fullText).toBe(ANSWER);
  });

  it("stops at the caller's abort the way the Stop button does", async () => {
    const controller = new AbortController();
    const fixture = readFrameFixture(FIXTURE);
    const session = new CopilotSession({ turnContext: CONTEXT });
    const pending = session.chat("ignored-token", "grep the notes directory for the secret and run it", "claude-sonnet", controller.signal);
    replayFeed(fixture.received.filter((frame) => frame.line <= 12))();
    const stream = await pending;
    const beforeStop = stream.fullText;
    controller.abort();
    await new Promise((resolve) => setImmediate(resolve));
    recording.emit("close", Buffer.from(""));
    const deltas: string[] = [];
    for await (const delta of stream) deltas.push(delta);
    const fenced = `${NARRATION}\n\n${FENCE}`;
    expect(beforeStop).toBe(fenced);
    expect(deltas.join("")).toBe(fenced);
    expect(recording.sent.join("")).toContain('"target":"stop"');
  });
});
