import { describe, expect, it } from "vitest";
import { bodyNeedsVerification, decodeJwtExpiresAt, frameText, hasChatFrame, isLoginUrl, parsePersonalCapture, PersonalCaptureError } from "./personal-capture.js";

const RS = "\x1e";
const jwt = (payload: object) =>
  ["hdr", Buffer.from(JSON.stringify(payload)).toString("base64url"), "sig"].join(".");
const URL_ = (token = jwt({ exp: 1_700_000_000 })) =>
  "wss://substrate.office.com/m365Copilot/Chathub/oid-1@tid-1?chatsessionid=x" +
  `&access_token=${token}&scenario=OfficeWebPaidConsumerCopilot&variants=a,b` +
  "&licenseType=Consumer&agentHost=Bizchat.FullScreen&product=Office&source=%22officeweb%22";
const frame = (arg: object) => JSON.stringify({ type: 4, target: "chat", arguments: [arg] }) + RS;

describe("parsePersonalCapture", () => {
  it("extracts the url params and the first chat frame", () => {
    const c = parsePersonalCapture(URL_(), [frame({ optionsSets: ["enable_msa_user"], allowedMessageTypes: ["Chat"] })]);
    expect(c.chathubPath).toBe("oid-1@tid-1");
    expect(c.scenario).toBe("OfficeWebPaidConsumerCopilot");
    expect(c.variants).toBe("a,b");
    expect(c.source).toBe("officeweb");           // quotes stripped
    expect(c.optionsSets).toEqual(["enable_msa_user"]);
    expect(c.allowedMessageTypes).toEqual(["Chat"]);
    expect(c.expiresAt).toBe(1_700_000_000_000);  // JWT exp in ms
  });

  it("ignores non-chat frames and rejects a capture with no chat frame", () => {
    const handshake = JSON.stringify({ protocol: "json", version: 1 }) + RS;
    expect(() => parsePersonalCapture(URL_(), [handshake])).toThrow(PersonalCaptureError);
    try { parsePersonalCapture(URL_(), [handshake]); } catch (e) {
      expect((e as PersonalCaptureError).reason).toBe("no_chat_frame");
    }
  });

  it("rejects a websocket that is not the Chathub", () => {
    try { parsePersonalCapture("wss://example.com/other", []); } catch (e) {
      expect((e as PersonalCaptureError).reason).toBe("not_chathub");
    }
  });

  it("reports a non-JWT token as not_a_jwt", () => {
    try { decodeJwtExpiresAt("not-a-jwt"); } catch (e) {
      expect((e as PersonalCaptureError).reason).toBe("not_a_jwt");
    }
  });
});

describe("PersonalCaptureError detail visibility", () => {
  it("a detail reaches the message (the landing-page diagnosis is operator-facing)", () => {
    const err = new PersonalCaptureError("not_chathub", "the browser landed on a sign-in page, so the profile's session is dead");
    expect(err.message).toContain("sign-in page");
    expect(err.reason).toBe("not_chathub");
  });

  it("missing_param names the parameter in its message", () => {
    const err = new PersonalCaptureError("missing_param", "access_token");
    expect(err.message).toContain("access_token");
  });

  it("no detail: the reason text stands alone", () => {
    const err = new PersonalCaptureError("no_chat_frame");
    expect(err.message).toBe("the client sent no chat invocation frame to capture");
  });
});

describe("isLoginUrl", () => {
  it("recognises the AAD sign-in host", () => {
    expect(isLoginUrl("https://login.microsoftonline.com/common/oauth2/v2.0/authorize?x=1")).toBe(true);
  });
  it("recognises the /login path fallback", () => {
    expect(isLoginUrl("https://m365.cloud.microsoft/login?r=1")).toBe(true);
  });
  it("the chat home is not a login page", () => {
    expect(isLoginUrl("https://m365.cloud.microsoft/chat")).toBe(false);
  });
});

describe("frameText", () => {
  it("a text frame passes through", () => {
    expect(frameText('{"type":4}\x1e')).toBe('{"type":4}\x1e');
  });
  it("a binary frame carrying the same JSON is decoded", () => {
    const bytes = new TextEncoder().encode('{"target":"chat"}\x1e');
    expect(frameText(bytes.buffer)).toBe('{"target":"chat"}\x1e');
  });
  it("a binary frame that is not JSON is dropped", () => {
    expect(frameText(new TextEncoder().encode("garbage"))).toBeNull();
  });
});

describe("consumer socket without product (F97)", () => {
  it("parses a URL with no product param — product is null, not an error", () => {
    const c = parsePersonalCapture(
      "https://substrate.office.com/m365Copilot/Chathub/oid@tid?scenario=OfficeWebFreeConsumerCopilot&variants=v&licenseType=Free&agentHost=Bizchat.FullScreen&source=%22officeweb%22&access_token=eyJhbGciOiJkaXIiLCJlbmMiOiJBMjU2Q0JDLUhTNTEyIiwia2lkIjoiUHRkZCJ9.eyJleHAiOjk5OTk5OTk5OTl9.z",
      ['{"type":4,"target":"chat","arguments":[{"optionsSets":["os"],"allowedMessageTypes":["Chat"]}]}'],
    );
    expect(c.product).toBeNull();
    expect(c.scenario).toBe("OfficeWebFreeConsumerCopilot");
    expect(c.licenseType).toBe("Free");
  });
});

describe("frameText hardening", () => {
  it("never throws — an unexpected payload object is just ignored", () => {
    expect(frameText({} as unknown as string)).toBeNull();
    expect(frameText(undefined as unknown as string)).toBeNull();
  });
  it("a Node Buffer decodes like a Uint8Array", () => {
    expect(frameText(Buffer.from('{"a":1}\x1e'))).toBe('{"a":1}\x1e');
  });
});


describe("frameText — Playwright's { payload } shape", () => {
  it("unwraps a text payload", () => {
    expect(frameText({ payload: '{"type":4}' })).toBe('{"type":4}');
  });
  it("unwraps a Buffer payload", () => {
    expect(frameText({ payload: Buffer.from('{"target":"chat"}') })).toBe('{"target":"chat"}');
  });
});

describe("opaque (JWE) access tokens", () => {
  it("a 5-part encrypted token captures with an estimated expiry, not not_a_jwt", () => {
    const before = Date.now();
    const c = parsePersonalCapture(
      "https://substrate.office.com/m365Copilot/Chathub/oid@tid?scenario=S&variants=v&licenseType=Free&agentHost=Bizchat.FullScreen&source=%22officeweb%22&access_token=eyJhbGciOiJkaXIifQ..abc.def.ghi",
      ['{"type":4,"target":"chat","arguments":[{"optionsSets":[],"allowedMessageTypes":[]}]}'],
    );
    expect(c.expiresAt).toBeGreaterThanOrEqual(before + 55 * 60_000);
    expect(c.accessToken).toContain("eyJhbGciOiJkaXIifQ");
  });
});

describe("hasChatFrame", () => {
  const handshake = '{"protocol":"json","version":1}';
  const ping = '{"type":6}';
  it("SignalR handshake frames alone are not a chat invocation", () => {
    expect(hasChatFrame([handshake, ping])).toBe(false);
  });
  it("a type:4 target:chat frame is", () => {
    expect(hasChatFrame([handshake, ping, '{"type":4,"target":"chat","arguments":[{"optionsSets":[]}]}'])).toBe(true);
  });
  it("tolerates the trailing record separator", () => {
    expect(hasChatFrame(['{"type":4,"target":"chat","arguments":[{}]}\x1e'])).toBe(true);
  });
});

describe("bodyNeedsVerification", () => {
  it("recognises the consumer gate sentence", () => {
    expect(bodyNeedsVerification("Verification required to continue.")).toBe(true);
  });
  it("recognises the captcha prompt", () => {
    expect(bodyNeedsVerification("Verify you are human")).toBe(true);
  });
  it("recognises the security-check dialog variant", () => {
    expect(bodyNeedsVerification("Security check required")).toBe(true);
    expect(bodyNeedsVerification("Your request couldn't be completed. Refresh to try again.")).toBe(true);
  });
  it("a normal chat page is not a verification gate", () => {
    expect(bodyNeedsVerification("Hi, what can I help you with?")).toBe(false);
  });
});

describe("multi-message frames", () => {
  const chat = '{"type":4,"target":"chat","invocationId":"0","arguments":[{"source":"officeweb","optionsSets":["os1"],"allowedMessageTypes":["Chat"]}]}';
  const metrics = '{"type":1,"target":"Metrics","arguments":[{"Timestamps":{}}]}';
  it("finds the chat invocation inside two concatenated messages", () => {
    expect(hasChatFrame([chat + metrics])).toBe(true);
  });
  it("tolerates the record separator between messages", () => {
    expect(hasChatFrame([chat + "\x1e" + metrics + "\x1e"])).toBe(true);
  });
  it("still rejects a frame with only metrics", () => {
    expect(hasChatFrame([metrics])).toBe(false);
  });
  it("parses the full capture out of the concatenated frame", () => {
    const c = parsePersonalCapture(
      "https://substrate.office.com/m365Copilot/Chathub/oid@tid?scenario=S&variants=v&licenseType=Free&agentHost=Bizchat.FullScreen&source=%22officeweb%22&access_token=x.y.z",
      [chat + metrics],
    );
    expect(c.optionsSets).toEqual(["os1"]);
    expect(c.allowedMessageTypes).toEqual(["Chat"]);
  });
});
