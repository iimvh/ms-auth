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

const fixturePath = (name: string) => fileURLToPath(new URL(`../../fixtures/replay/${name}`, import.meta.url));

const CONTEXT: TurnContext = {
  token: "jwe-token",
  socketUrl: "wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?scenario=S",
  headers: { Origin: "https://m365.cloud.microsoft", "User-Agent": "replay-agent" },
  payloadDefaults: { scenario: "OfficeWebFreeConsumerCopilot", licenseType: "Free" },
  capabilities: ["code_interpreter"],
  accountId: "personal",
  personal: true,
  record: null,
};

const PLAIN = fixturePath("recorded-plain-answer.ndjson");
const THROTTLE = fixturePath("recorded-throttling-update.ndjson");
const DISENGAGED = fixturePath("recorded-disengaged.ndjson");

const PLAIN_ANSWER = "Paris is the capital of France.";
const REFUSAL_ANSWER = "Sorry, it looks like I can't chat about this. Let's try a different topic.";

async function replay(fixtureFile: string) {
  const fixture = readFrameFixture(fixtureFile);
  const stream = await replayFrames(fixture);
  const deltas: string[] = [];
  for await (const delta of stream) deltas.push(delta);
  return { fixture, stream, streamed: deltas.join("") };
}

async function replayFrames(fixture: ReturnType<typeof readFrameFixture>) {
  const session = new CopilotSession({ turnContext: () => CONTEXT, accountId: "personal" });
  const pending = session.chat("ignored-token", "any user prompt", "claude-sonnet");
  replayFeed(fixture.received)();
  const stream = await pending;
  recording.emit("close", Buffer.from(""));
  return stream;
}

describe("recordings replayed as their own fixtures (#3)", () => {
  it.each([PLAIN, THROTTLE, DISENGAGED])("loads a real M365 capture with no synthetic note (%s)", (file) => {
    const fixture = readFrameFixture(file);
    expect(fixture.synthetic).toBeNull();
    expect(fixture.lines.length).toBeGreaterThan(3);
    expect(fixture.lines.some((line) => line.dir === "send")).toBe(true);
    expect(fixture.received.length).toBeGreaterThan(3);
    expect(fixture.lines.filter((line) => (line.frame as { type?: number }).type === 3)).toHaveLength(1);
  });

  it("replays the recorded plain answer, folding it out of cursor deltas", async () => {
    const { fixture, stream, streamed } = await replay(PLAIN);
    expect(stream.fullText).toBe(PLAIN_ANSWER);
    expect(streamed).toBe(PLAIN_ANSWER);
    const updates = fixture.received.map((frame) => JSON.parse(frame.data.toString().slice(0, -1)));
    const head = updates.find((frame) => frame.arguments?.[0]?.messages?.some?.((m: any) => m.text === "Paris"));
    expect(head.arguments[0].cursor).toMatchObject({ p: -1 });
    expect(head.arguments[0].cursor.j).toContain("adaptiveCards[0].body[0].text");
    const delta = updates.filter((frame) => frame.arguments?.[0]?.writeAtCursor);
    expect(delta.map((frame) => frame.arguments[0].writeAtCursor).join("")).toBe(" is the capital of France.");
    expect((head.arguments[0].messages.find((m: any) => m.text === "Paris") as any).text + delta.map((frame) => frame.arguments[0].writeAtCursor).join("")).toBe(PLAIN_ANSWER);
    expect(stream.throttle).toEqual({ current: 1, max: 600 });
    expect(stream.turnState).toBe("Completed");
    expect(stream.turnCount).toBe(1);
  });

  it("replays the recorded throttling update exactly as M365 sent it", async () => {
    const { stream } = await replay(THROTTLE);
    expect(stream.throttle).toEqual({ current: 1, max: 600 });
    const throttles = readFrameFixture(THROTTLE).lines
      .filter((line) => line.dir === "recv")
      .map((line) => line.frame as any)
      .filter((frame: any) => frame.arguments?.some?.((arg: any) => arg.throttling) || frame.item?.throttling);
    expect(throttles.length).toBeGreaterThanOrEqual(2);
    for (const frame of throttles) {
      const t = frame.item?.throttling ?? frame.arguments.find((arg: any) => arg.throttling).throttling;
      expect(t.numUserMessagesInConversation).toBe(1);
      expect(t.maxNumUserMessagesInConversation).toBe(600);
      expect(t.numLongDocSummaryUserMessagesInConversation).toBe(0);
    }
  });

  it("replays the recorded Disengaged turn: an apology plus the marker, no answer", async () => {
    const { fixture, stream, streamed } = await replay(DISENGAGED);
    expect(streamed).toBe(REFUSAL_ANSWER);
    const disengaged = fixture.lines.filter(
      (line) => line.dir === "recv" && JSON.stringify(line.frame).includes('"messageType":"Disengaged"'),
    );
    expect(disengaged.length).toBeGreaterThanOrEqual(2);
    const updates = disengaged.map((line) => line.frame).filter((frame: any) => frame.type === 1 && frame.target === "update");
    expect(updates.length).toBe(1);
    const markerMessages = [
      ...updates.flatMap((frame: any) => frame.arguments.flatMap((arg: any) => arg.messages ?? [])),
      ...disengaged.map((line) => line.frame).filter((frame: any) => frame.type === 2).flatMap((frame: any) => frame.item.messages),
    ].filter((m: any) => m.messageType === "Disengaged");
    expect(markerMessages.length).toBe(3);
    for (const marker of markerMessages) {
      expect(marker.author).toBe("bot");
      expect(["None", "Unknown"]).toContain(marker.offense);
      expect(marker.text).toBeUndefined();
      expect(marker.hiddenText).toMatch(/Disengaged|disengaged/);
    }
    expect(markerMessages).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ contentOrigin: "JailBreakClassifier", hiddenText: "Disengaged by JailBreak Classifier" }),
        expect.objectContaining({ contentOrigin: "DeepLeo", hiddenText: "> Conversation disengaged" }),
      ]),
    );
    expect(stream.turnState).toBe("Completed");
    expect(stream.throttle).toEqual({ current: 1, max: 600 });
    expect(stream.messageType).toBe("Disengaged");
    expect(stream.result?.value).toBe("Success");
  });
});
