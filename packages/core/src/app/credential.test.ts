import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

vi.mock("../adapters/personal/personal-harvest-cdp.js", () => ({ attachChatPage: () => Promise.reject(new Error("no browser")) }));
vi.mock("../adapters/msal/auth.js", () => ({ getToken: async () => "enterprise-token" }));

vi.mock("ws", async () => ({ default: (await import("../protocol/socket-fixture.js")).RecordingWebSocket }));

const { recording, handshake, botText } = await import("../protocol/socket-fixture.js");
const { accountsDir, activateAccount, clearMockRegistry, useMockRegistry } = await import("../account-registry.js");
const { PersonalStore } = await import("../adapters/personal/personal-session.js");
const { selectCredential, turnContextFor } = await import("./credential.js");
const { ModelSession } = await import("./model.js");

const ANSWER = "an answer from the chosen account";

function record(token: string) {
  return {
    accessToken: token,
    chathubPath: "oid@tid",
    scenario: "OfficeWebFreeConsumerCopilot",
    variants: "v1,v2",
    licenseType: "Free",
    agentHost: "Bizchat.FullScreen",
    product: null,
    source: "officeweb",
    optionsSets: ["cwc_code_interpreter"],
    allowedMessageTypes: ["Chat"],
    emailHint: null,
    probe: null,
    mintCount: 1,
    lastMintAt: 0,
    expiresAt: Date.now() + 3_600_000,
    capturedAt: Date.now() - 60_000,
  };
}

function twoAccounts() {
  process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "cred-home-"));
  useMockRegistry([{ id: "default", hasSession: true }, { id: "work", hasSession: true }]);
  new PersonalStore(join(accountsDir(), "default", "session.json")).write(record("token-default"));
  new PersonalStore(join(accountsDir(), "work", "session.json")).write(record("token-work"));
}

/** Run one real turn through ModelSession and the real transport. */
async function turn(accountId: string | null): Promise<{ url: string; payload: string; text: string }> {
  const session = new ModelSession({ accountId, useAgent: false });
  const pending = session.run("hi", "m365-copilot", undefined, false);
  await new Promise((resolve) => setImmediate(resolve));
  recording.emit("message", handshake());
  recording.emit("message", botText(ANSWER));
  const stream = await pending;
  recording.emit("close", Buffer.from(""));
  let text = "";
  for await (const delta of stream) text += delta;
  return { url: recording.url, payload: recording.sent.join(""), text };
}

beforeEach(() => {
  process.env.M365_AUTH = "personal";
});

afterEach(() => {
  clearMockRegistry();
  delete process.env.M365_AUTH;
  delete process.env.M365_ACCOUNT_HOME;
  delete process.env.M365_PERSONAL_SESSION_FILE;
  delete process.env.M365_SCENARIO;
});

describe("the one credential decision (#9)", () => {
  it("returns the pinned account's token, record and account id", async () => {
    twoAccounts();
    activateAccount("default");
    const selection = await selectCredential("work");
    expect(selection.accountId).toBe("work");
    expect(selection.personal).toBe(true);
    expect(selection.token).toBe("token-work");
    expect(selection.record?.accessToken).toBe("token-work");
  });

  it("returns the active account when no account is pinned", async () => {
    twoAccounts();
    activateAccount("default");
    expect((await selectCredential()).accountId).toBe("default");
    expect((await selectCredential(null)).token).toBe("token-default");
    activateAccount("work");
    expect((await selectCredential()).accountId).toBe("work");
    expect((await selectCredential(null)).token).toBe("token-work");
  });

  it("keeps a pinned account after the active one switches away", async () => {
    twoAccounts();
    activateAccount("default");
    activateAccount("work");
    const selection = await selectCredential("default");
    expect(selection.accountId).toBe("default");
    expect(selection.token).toBe("token-default");
    expect(selection.record?.chathubPath).toBe("oid@tid");
  });

  it("is not personal outside personal mode: no account, no record", async () => {
    delete process.env.M365_AUTH;
    const selection = await selectCredential("work");
    expect(selection).toMatchObject({ accountId: null, personal: false, record: null });
    expect(selection.token).toBe("enterprise-token");
  });
});

describe("a turn on a pinned account (#9)", () => {
  it("runs on that account's token, chathub path and payload, not the active one's", async () => {
    twoAccounts();
    activateAccount("work");
    const { url, payload, text } = await turn("default");

    expect(text).toBe(ANSWER);
    expect(url).toContain("access_token=token-default");
    expect(url).not.toContain("token-work");
    expect(payload).toContain('"tone":"Magic"');
    expect(payload).toContain('"ProductCategory":"Chat"');
    expect(payload).toContain('"deviceOS":"Windows"');
    expect(payload).toContain("cwc_code_interpreter");
  });

  it("stays on the pinned account after a mid-thread switch of the active account", async () => {
    twoAccounts();
    activateAccount("default");
    const session = new ModelSession({ accountId: "default", useAgent: false });
    activateAccount("work");
    const pending = session.run("hi", "m365-copilot", undefined, false);
    await new Promise((resolve) => setImmediate(resolve));
    recording.emit("message", handshake());
    await pending;
    expect(recording.url).toContain("access_token=token-default");
  });
});

describe("a turn on the active account (#9)", () => {
  it("runs on whichever account is active, and follows a switch", async () => {
    twoAccounts();
    activateAccount("default");
    const first = await turn(null);
    expect(first.text).toBe(ANSWER);
    expect(first.url).toContain("access_token=token-default");

    activateAccount("work");
    const second = await turn(null);
    expect(second.text).toBe(ANSWER);
    expect(second.url).toContain("access_token=token-work");
  });

  it("builds the builtin layout's record path when the registry has no accounts", async () => {
    process.env.M365_PERSONAL_SESSION_FILE = join(mkdtempSync(join(tmpdir(), "cred-builtin-")), "session.json");
    new PersonalStore(process.env.M365_PERSONAL_SESSION_FILE).write(record("token-builtin"));
    const selection = await selectCredential();
    expect(selection.accountId).toBeNull();
    expect(selection.token).toBe("token-builtin");
  });
});

describe("the context the transport is given", () => {
  it("carries the account, the personal marker and no consumer-specific default leaks to the enterprise path", async () => {
    twoAccounts();
    const selection = await selectCredential("work");
    const plan = { requestId: "r", sessionId: "s", conversationId: "c", model: "m365-copilot", temporaryChat: true, agentId: undefined, nativeActions: false, images: false };
    const context = turnContextFor(selection, plan);
    expect(context).toMatchObject({ accountId: "work", personal: true, token: "token-work" });
    expect(context.socketUrl).toContain("/Chathub/oid@tid?");
    expect(context.socketUrl).toContain("ConversationId=c");
    expect(context.socketUrl).toContain("disableMemory=1");
    expect(context.headers["User-Agent"]).toContain("Chrome/");
    expect(context.payloadDefaults).toMatchObject({ tone: "Magic", deviceOS: "Windows" });
  });

  it("adds the image message type only when the turn asks for images", async () => {
    twoAccounts();
    const selection = await selectCredential("work");
    const base = { requestId: "r", sessionId: "s", conversationId: "c", model: "m365-copilot", temporaryChat: true, agentId: undefined, nativeActions: false };
    const plain = turnContextFor(selection, { ...base, images: false });
    const withImages = turnContextFor(selection, { ...base, images: true });
    expect((plain.payloadDefaults.allowedMessageTypes as string[])).not.toContain("GenerateGraphicArt");
    expect((withImages.payloadDefaults.allowedMessageTypes as string[])).toContain("GenerateGraphicArt");
  });
});
