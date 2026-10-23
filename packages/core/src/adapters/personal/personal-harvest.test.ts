import { describe, expect, it, vi } from "vitest";
import {
  harvestPersonalToken,
  HarvestError,
  tokenFromSocketUrl,
  type HarvestedPage,
  type HarvestedSocket,
} from "./personal-harvest.js";
import type { PersonalCapture } from "./personal-session.js";

const CHATHUB =
  "wss://substrate.office.com/m365Copilot/Chathub/oid@tid?scenario=OfficeWebFreeConsumerCopilot" +
  "&variants=v&licenseType=Free&agentHost=Bizchat.FullScreen&source=%22officeweb%22&access_token=TOKEN";

const prior: PersonalCapture = {
  accessToken: "OLD",
  chathubPath: "oid@tid",
  scenario: "OfficeWebFreeConsumerCopilot",
  variants: "v",
  licenseType: "Free",
  agentHost: "Bizchat.FullScreen",
  product: null,
  source: "officeweb",
  optionsSets: ["a"],
  allowedMessageTypes: ["Chat"],
  expiresAt: 1,
  capturedAt: 0,
};

/** A page handle that reports the socket a harvest should find. */
function pageThatOpens(socket: HarvestedSocket | null): HarvestedPage & { closed: boolean } {
  const page = {
    closed: false,
    url: () => "https://m365.cloud.microsoft/chat",
    sendAndWait: vi.fn(async () => socket),
    close: vi.fn(async () => {
      page.closed = true;
    }),
  };
  return page as HarvestedPage & { closed: boolean };
}

const run = (over: Partial<Parameters<typeof harvestPersonalToken>[0]> = {}) => {
  const page = pageThatOpens(over.attach ? null : { url: CHATHUB, frames: ['{"target":"chat"}'] });
  return {
    page,
    promise: harvestPersonalToken({
      attach: async () => page,
      previousToken: "OLD",
      waitMs: 1,
      now: () => 5_000,
      ...over,
    }),
  };
};

describe("tokenFromSocketUrl", () => {
  it("reads access_token out of a Chathub URL", () => {
    expect(tokenFromSocketUrl(CHATHUB)).toBe("TOKEN");
  });

  it("is null for a URL without one or a non-URL", () => {
    expect(tokenFromSocketUrl("wss://substrate.office.com/x")).toBeNull();
    expect(tokenFromSocketUrl("not a url")).toBeNull();
  });
});

describe("harvestPersonalToken", () => {
  it("reads the token off the socket the send opened", async () => {
    const { promise } = run();
    const result = await promise;
    expect(result.capture.accessToken).toBe("TOKEN");
    expect(result.spentTurn).toBe(true);
  });

  it("reports a rotated token when it differs from what we had", async () => {
    const { promise } = run({ previousToken: "OLD" });
    expect((await promise).rotated).toBe(true);
  });

  it("reports no rotation when the page reused its token — and still spent a turn", async () => {
    // Seen once (two sends 3 min apart, F111); 5/6 rounds 10 min apart rotated.
    const { promise } = run({ previousToken: "TOKEN" });
    const result = await promise;
    expect(result.rotated).toBe(false);
    expect(result.spentTurn).toBe(true);
  });

  it("rebuilds the connection params from the URL, keeping the observed values", async () => {
    const { promise } = run();
    const { capture } = await promise;
    expect(capture.scenario).toBe("OfficeWebFreeConsumerCopilot");
    expect(capture.licenseType).toBe("Free");
    expect(capture.chathubPath).toBe("oid@tid");
    expect(capture.product).toBeNull();
    expect(capture.capturedAt).toBe(5_000);
  });

  it("reuses the caller's params instead of re-deriving them", async () => {
    const { promise } = run({
      reuse: (socketUrl) => ({ ...prior, accessToken: tokenFromSocketUrl(socketUrl) ?? "" }),
    });
    const { capture } = await promise;
    expect(capture.optionsSets).toEqual(["a"]);
    expect(capture.accessToken).toBe("TOKEN");
  });

  it("no socket after the send is no_chat_frame", async () => {
    const page = pageThatOpens(null);
    const err = await harvestPersonalToken({
      attach: async () => page,
      previousToken: "OLD",
      waitMs: 1,
      now: () => 0,
    }).catch((e) => e);
    expect(err).toBeInstanceOf(HarvestError);
    expect(err.reason).toBe("no_chat_frame");
  });

  it("a non-Chathub socket is capture_failed", async () => {
    const page = pageThatOpens({ url: "wss://go.trouter.skype.com/v4/c", frames: ['{"target":"chat"}'] });
    const err = await harvestPersonalToken({
      attach: async () => page,
      previousToken: "OLD",
      waitMs: 1,
      now: () => 0,
    }).catch((e) => e);
    expect(err.reason).toBe("capture_failed");
  });

  it("a Chathub socket with no token is capture_failed", async () => {
    const page = pageThatOpens({
      url: "wss://substrate.office.com/m365Copilot/Chathub/oid@tid?scenario=s&variants=v&licenseType=l&agentHost=a&source=s",
      frames: ['{"target":"chat"}'],
    });
    const err = await harvestPersonalToken({
      attach: async () => page,
      previousToken: "OLD",
      waitMs: 1,
      now: () => 0,
    }).catch((e) => e);
    expect(err.reason).toBe("capture_failed");
  });

  it("always closes the page handle, even when the send yields nothing", async () => {
    const page = pageThatOpens(null);
    await harvestPersonalToken({ attach: async () => page, previousToken: "OLD", waitMs: 1, now: () => 0 }).catch(
      () => {},
    );
    expect(page.closed).toBe(true);
  });

  it("an unreachable browser propagates as chrome_unreachable", async () => {
    const err = await harvestPersonalToken({
      attach: async () => {
        throw new HarvestError("chrome_unreachable", "ECONNREFUSED");
      },
      previousToken: "OLD",
      waitMs: 1,
      now: () => 0,
    }).catch((e) => e);
    expect(err.reason).toBe("chrome_unreachable");
  });
});
