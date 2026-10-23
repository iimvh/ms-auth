import { describe, it, expect, afterEach } from "vitest";
import { cursorMessageId, foldStreamText, TurnTextComposer } from "./session.js";
import { buildCopilotWebSocketHeaders, buildCopilotWebSocketUrl, buildPersonalChathubUrl, buildPersonalWebSocketHeaders, personalTurnParams, readPersonalRecord } from "../personal-chathub/chathub.js";
import { accountsDir, activateAccount, clearMockRegistry, useMockRegistry } from "../account-registry.js";
import { PersonalStore } from "../adapters/personal/personal-session.js";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { MessageUpdate } from "./schemas.js";

/** Replay a sequence of raw M365 frames (deltas as {d}, snapshots as {s}) through
 *  foldStreamText and collect what would be streamed + the final buffered answer. */
function replay(frames: Array<{ d: string } | { s: string }>) {
  let answer = "";
  const emitted: string[] = [];
  for (const f of frames) {
    const next = "d" in f ? answer + f.d : f.s;
    const r = foldStreamText(answer, next);
    answer = r.answer;
    if (r.emit) emitted.push(r.emit);
  }
  return { emitted, streamed: emitted.join(""), answer };
}

describe("foldStreamText", () => {
  it("streams pure token deltas as-is", () => {
    const r = replay([{ d: "Hello" }, { d: ", " }, { d: "world" }]);
    expect(r.emitted).toEqual(["Hello", ", ", "world"]);
    expect(r.answer).toBe("Hello, world");
  });

  it("recovers the head token when it arrives only as a snapshot (the live bug)", () => {
    // M365 delivered "alpha" as a full-text snapshot, then token deltas for the rest.
    const r = replay([
      { s: "alpha" },
      { d: "\nbeta" },
      { d: "\ngamma" },
    ]);
    expect(r.streamed).toBe("alpha\nbeta\ngamma");
    expect(r.answer).toBe("alpha\nbeta\ngamma");
  });

  it("streams the appended suffix when snapshots grow monotonically", () => {
    const r = replay([{ s: "The " }, { s: "The quick " }, { s: "The quick fox" }]);
    expect(r.emitted).toEqual(["The ", "quick ", "fox"]);
    expect(r.answer).toBe("The quick fox");
  });

  it("emits the extending tail of a snapshot exactly once (no duplication)", () => {
    // Deltas gave "Hello wor"; a final snapshot "Hello world" EXTENDS it, so only the
    // "ld" tail is emitted — never the whole snapshot on top of the deltas.
    const r = replay([{ d: "Hello " }, { d: "wor" }, { s: "Hello world" }]);
    expect(r.emitted).toEqual(["Hello ", "wor", "ld"]);
    expect(r.streamed).toBe("Hello world");
    expect(r.answer).toBe("Hello world");
  });

  it("ignores shorter/equal snapshots (no negative-length slice, no re-emit)", () => {
    const r = replay([{ d: "abcdef" }, { s: "abc" }, { s: "abcdef" }]);
    expect(r.emitted).toEqual(["abcdef"]);
    expect(r.answer).toBe("abcdef");
  });

  it("adopts a divergent longer snapshot as authoritative but does not stream it", () => {
    // Deltas streamed "xy"; a later snapshot reveals a different, longer head. We keep
    // the snapshot for the buffered result but must NOT stream it (can't unsend "xy").
    const r = replay([{ d: "xy" }, { s: "ZZZxy extra" }]);
    expect(r.streamed).toBe("xy");
    expect(r.answer).toBe("ZZZxy extra");
    expect(r.answer.startsWith(r.streamed)).toBe(false); // divergence recorded, not streamed
  });
});

describe("GraphicArt image frame parsing (§14)", () => {
  // The exact shape captured from a live GUI image turn — the fields zod used to
  // strip. If BotMessage is ever re-tightened, this fails and the image is lost.
  const graphicArtFrame = {
    messages: [
      {
        text: "Loading image",
        contentGenerationProgressList: [
          {
            contentType: "image",
            size: "Xlimage",
            orientation: "Landscape",
            pollUrl: "eyJQb2xsSWQiOiJ4In0=",
            fileToken: "359965a5-6ca9-409d-a4ca-43b0cc9cdf81",
            ImageReferenceUrls: ["https://designerapp.officeapps.live.com/designerapp/document.ashx?path=%2Fx%2FDallEGeneratedImages%2Fdalle-abc.png"],
            status: 2,
          },
        ],
        contentType: "GraphicArt",
        author: "bot",
        messageType: "Progress",
        contentOrigin: "ImageGeneration",
      },
    ],
  };

  it("retains the image payload instead of stripping it", () => {
    const parsed = MessageUpdate.safeParse(graphicArtFrame);
    expect(parsed.success).toBe(true);
    const m = parsed.data!.messages[0] as any;
    expect(m.contentType).toBe("GraphicArt");
    const entry = m.contentGenerationProgressList[0];
    expect(entry.ImageReferenceUrls[0]).toContain("DallEGeneratedImages");
    expect(entry.fileToken).toBe("359965a5-6ca9-409d-a4ca-43b0cc9cdf81");
    expect(entry.status).toBe(2);
  });

  it("still parses an ordinary text bot message with no image fields", () => {
    const parsed = MessageUpdate.safeParse({
      messages: [{ text: "just text", author: "bot", messageType: "Chat" }],
    });
    expect(parsed.success).toBe(true);
    const m = parsed.data!.messages[0] as any;
    expect(m.contentGenerationProgressList).toBeUndefined();
  });
});

describe("temporary-chat WebSocket URL", () => {
  it("preserves disableMemory=1 when temporary chat is enabled", () => {
    const params = new URLSearchParams({
      ConversationId: "conversation-1",
      disableMemory: "1",
    });
    const url = new URL(buildCopilotWebSocketUrl("oid-1", "tid-1", params));

    expect(url.hostname).toBe("substrate.office.com");
    expect(url.pathname).toBe("/m365Copilot/Chathub/oid-1@tid-1");
    expect(url.searchParams.get("ConversationId")).toBe("conversation-1");
    expect(url.searchParams.get("disableMemory")).toBe("1");
  });

  it("does not invent disableMemory when saved history is requested", () => {
    const params = new URLSearchParams({ ConversationId: "conversation-1" });
    const url = new URL(buildCopilotWebSocketUrl("oid-1", "tid-1", params));

    expect(url.searchParams.has("disableMemory")).toBe(false);
  });
});

describe("Chathub WebSocket headers", () => {
  afterEach(() => { delete process.env.M365_NO_SESSION_ROUTING; });

  it("pins the conversation to one backend by its ConversationId", () => {
    expect(buildCopilotWebSocketHeaders("conversation-1")["X-RoutingParameter-SessionKey"]).toBe("conversation-1");
  });

  it("leaves the routing key out under M365_NO_SESSION_ROUTING", () => {
    process.env.M365_NO_SESSION_ROUTING = "1";
    expect(buildCopilotWebSocketHeaders("conversation-1")).not.toHaveProperty("X-RoutingParameter-SessionKey");
  });
});

describe("TurnTextComposer (multi-message turns)", () => {
  type F = { cursor?: string } & ({ d: string } | { s: string; id: string });
  /** Replay frames through the composer AND the streaming fold, like session.ts. */
  function compose(frames: F[]) {
    const c = new TurnTextComposer();
    let answer = "";
    let streamed = "";
    for (const f of frames) {
      if (f.cursor) c.cursor(cursorMessageId(f.cursor)!); // as session.ts does
      if ("d" in f) c.delta(f.d);
      else c.snapshot(f.id, f.s);
      const r = foldStreamText(answer, c.text);
      answer = r.answer;
      if (r.emit) streamed += r.emit;
    }
    return { text: c.text, answer, streamed };
  }
  const cur = (id: string) => `$['${id}'].adaptiveCards[0].body[0].text`;

  it("keeps the head of a second message — the live turn that lost a tool fence", () => {
    // Verbatim shape of the Sep 28 Sonnet 4.6 find-needle turn: narration in one
    // message, the ```bash fence in the next. The old single-string fold dropped
    // the second message's head snapshot "```" and produced "…SECRET_CODE.bash\ngrep".
    const r = compose([
      { cursor: cur("m1"), s: "Let", id: "m1" },
      { d: " me look" }, { d: " through" }, { d: " the files" }, { d: " in the notes/" },
      { d: " directory to find the SECRET_CODE" }, { d: "." },
      { s: "Let me look through the files in the notes/ directory to find the SECRET_CODE.", id: "m1" },
      { cursor: cur("m2"), s: "```", id: "m2" },
      { d: "bash\ngrep -r" }, { d: ' "^' }, { d: 'SECRET_CODE=" notes' }, { d: "/" }, { d: "\n```" },
      { s: '```bash\ngrep -r "^SECRET_CODE=" notes/\n```', id: "m2" },
    ]);
    expect(r.text).toBe(
      'Let me look through the files in the notes/ directory to find the SECRET_CODE.\n\n```bash\ngrep -r "^SECRET_CODE=" notes/\n```',
    );
    expect(r.answer).toBe(r.text);
    expect(r.streamed).toBe(r.text); // prefix-safe: everything streamed, nothing duplicated
  });

  it("is byte-identical to the old fold for a single-message turn", () => {
    const frames: F[] = [{ cursor: cur("a"), s: "alpha", id: "a" }, { d: "\nbeta" }, { s: "alpha\nbeta", id: "a" }];
    expect(compose(frames).text).toBe("alpha\nbeta");
    expect(compose(frames).streamed).toBe("alpha\nbeta");
  });

  it("routes deltas by cursor, not by whichever snapshot arrived last", () => {
    // A late final snapshot of message 1 lands after message 2's cursor; the
    // deltas that follow still belong to message 2.
    const r = compose([
      { cursor: cur("m1"), s: "one", id: "m1" },
      { cursor: cur("m2"), s: "two", id: "m2" },
      { s: "one!", id: "m1" },
      { d: " more" },
    ]);
    expect(r.text).toBe("one!\n\ntwo more");
  });

  it("falls back to the snapshot's message when no cursor was ever sent", () => {
    const r = compose([{ s: "x", id: "m1" }, { d: "yz" }]);
    expect(r.text).toBe("xyz");
  });

  it("parses the message id out of a cursor path", () => {
    expect(cursorMessageId("$['1051ab91-f905'].adaptiveCards[0].body[0].text")).toBe("1051ab91-f905");
    expect(cursorMessageId(undefined)).toBeNull();
    expect(cursorMessageId("garbage")).toBeNull();
  });
});

import type { PersonalCapture } from "../adapters/personal/personal-session.js";

const CAPTURE: PersonalCapture = {
  accessToken: "tok", chathubPath: "oid-1@tid-1", scenario: "OfficeWebPaidConsumerCopilot",
  variants: "a,b", licenseType: "Consumer", agentHost: "Bizchat.FullScreen", product: "Office",
  source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"], expiresAt: 1, capturedAt: 0,
};

describe("personal Chathub URL", () => {
  it("personalTurnParams uses the captured values, not the scenario table", () => {
    const p = personalTurnParams(CAPTURE, { requestId: "r", sessionId: "s", conversationId: "c" });
    expect(p.get("scenario")).toBe("OfficeWebPaidConsumerCopilot");
    expect(p.get("access_token")).toBe("tok");
    expect(p.get("variants")).toBe("a,b");
    expect(p.get("licenseType")).toBe("Consumer");
    expect(p.get("agentHost")).toBe("Bizchat.FullScreen");
    expect(p.get("chatsessionid")).toBe("r");
    expect(p.get("ConversationId")).toBe("c");
  });

  it("buildPersonalChathubUrl targets the captured path", () => {
    const url = buildPersonalChathubUrl(CAPTURE, new URLSearchParams({ scenario: "S" }));
    expect(url.startsWith("wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?")).toBe(true);
    expect(url).toContain("scenario=S");
  });
});

describe("personal M365_SCENARIO override", () => {
  afterEach(() => { delete process.env.M365_SCENARIO; });

  it("the operator override wins over the captured scenario", () => {
    process.env.M365_SCENARIO = "OfficeWebIncludedCopilot";
    const p = personalTurnParams(CAPTURE, { requestId: "r", sessionId: "s", conversationId: "c" });
    expect(p.get("scenario")).toBe("OfficeWebIncludedCopilot");
  });

  it("unset, the captured scenario stands", () => {
    const p = personalTurnParams(CAPTURE, { requestId: "r", sessionId: "s", conversationId: "c" });
    expect(p.get("scenario")).toBe("OfficeWebPaidConsumerCopilot");
  });
});

describe("personal replay omits absent params (F97)", () => {
  it("a consumer capture (product null) rebuilds the URL without product", () => {
    const cap = { ...CAPTURE, product: null, scenario: "OfficeWebFreeConsumerCopilot", licenseType: "Free" };
    const p = personalTurnParams(cap, { requestId: "r", sessionId: "s", conversationId: "c" });
    expect(p.get("product")).toBeNull();
    expect(p.get("scenario")).toBe("OfficeWebFreeConsumerCopilot");
  });
});

describe("personal websocket headers (F103)", () => {
  it("looks like the real Windows Chrome consumer client", () => {
    const h = buildPersonalWebSocketHeaders("conv-1");
    expect(h.Origin).toBe("https://m365.cloud.microsoft");
    expect(h["User-Agent"]).toContain("Windows NT 10.0");
    expect(h["User-Agent"]).toContain("Chrome/");
    expect(h["Cache-Control"]).toBe("no-cache");
    expect(h.Pragma).toBe("no-cache");
    expect(h["Accept-Encoding"]).toContain("br");
    expect(h["Sec-WebSocket-Extensions"]).toBeUndefined();
    expect(h["X-RoutingParameter-SessionKey"]).toBe("conv-1");
  });
  it("the enterprise headers still say Firefox (unchanged)", () => {
    expect(buildCopilotWebSocketHeaders("c")["User-Agent"]).toContain("Firefox");
  });
});

describe("a personal failure names the account (Task 8)", () => {
  afterEach(() => { clearMockRegistry(); delete process.env.M365_ACCOUNT_HOME; });

  it("a missing record names the account and its own login command", () => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work" }]);
    activateAccount("work");
    let err: { code?: string; message: string } | null = null;
    try { readPersonalRecord(); } catch (e) { err = e as { code?: string; message: string }; }
    expect(err).not.toBeNull();
    expect(err!.message).toContain('account "work"');
    expect(err!.message).toContain("pnpm m365-personal login --account work");
    expect(err!.message).toContain("no session record");
  });

  it("a corrupt record says so instead of a flat session_dead", () => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work" }]);
    activateAccount("work");
    const f = join(accountsDir(), "work", "session.json");
    writeFileSync(f, "{ not json"); // a record IS there, but cannot be used
    let msg = "";
    try { readPersonalRecord(); } catch (e) { msg = (e as Error).message; }
    expect(msg).toContain('account "work"');
    expect(msg).toContain("unusable");
    expect(msg).toContain("invalid_json");
  });

  it("an expired record names the account", () => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work" }]);
    activateAccount("work");
    new PersonalStore(join(accountsDir(), "work", "session.json")).write({
      accessToken: "t", chathubPath: "o@t", scenario: "S", variants: "v", licenseType: "Free",
      agentHost: "h", product: null, source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
      emailHint: null, probe: null, mintCount: 1, lastMintAt: 0, expiresAt: Date.now() - 1000,
      capturedAt: Date.now() - 7_200_000 });
    let msg = "";
    try { readPersonalRecord(); } catch (e) { msg = (e as Error).message; }
    expect(msg).toContain('account "work"');
    expect(msg).toContain("expired");
  });

  // The bug this pins (whole-branch review): the thread's token came from the
  // ACTIVE account while its record was read from whatever was active now, so
  // a switch mid-thread sent account A's record with account B's token.
  it("a pinned account reads ITS record after a switch, not the newly active one's", () => {
    process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "ar-home-"));
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work", hasSession: true }]);
    const fresh = (token: string) => ({
      accessToken: token, chathubPath: "o@t", scenario: "S", variants: "v", licenseType: "Free",
      agentHost: "h", product: null, source: "officeweb" as const, optionsSets: [], allowedMessageTypes: ["Chat"],
      emailHint: null, probe: null, mintCount: 1, lastMintAt: 0, expiresAt: Date.now() + 3_600_000,
      capturedAt: Date.now() - 60_000 });
    new PersonalStore(join(accountsDir(), "default", "session.json")).write(fresh("token-default"));
    new PersonalStore(join(accountsDir(), "work", "session.json")).write(fresh("token-work"));
    activateAccount("default");
    // The thread is opened on default and the active account switches away.
    expect(readPersonalRecord("default").accessToken).toBe("token-default");
    activateAccount("work");
    expect(readPersonalRecord("default").accessToken).toBe("token-default"); // still its own
    expect(readPersonalRecord().accessToken).toBe("token-work"); // active has moved on
    expect(() => readPersonalRecord("nope")).toThrow(/does not exist/);
  });
});