import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import { PersonalStore } from "./personal-session.js";
import { personalProfileDir } from "./personal-session.js";
import { getPersonalToken } from "./personal-auth.js";

vi.mock("./personal-harvest.js", () => {
  class HarvestError extends Error {
    constructor(readonly reason: string, message: string) {
      super(message);
    }
  }
  return { HarvestError, isHarvestError: (err: unknown) => err instanceof HarvestError };
});

vi.mock("./personal-capture.js", () => ({
  PersonalCaptureError: class extends Error {},
  capturePersonalSession: vi.fn(async () => ({
    accessToken: "tok-minted",
    chathubPath: "oid@tid",
    scenario: "S",
    variants: "v",
    licenseType: "Consumer",
    agentHost: "Bizchat.FullScreen",
    product: "Office",
    source: "officeweb",
    optionsSets: [],
    allowedMessageTypes: ["Chat"],
    expiresAt: Date.now() + 3_600_000,
    capturedAt: Date.now(),
  })),
}));

const { capturePersonalSession } = await import("./personal-capture.js");

beforeEach(() => {
  (capturePersonalSession as ReturnType<typeof vi.fn>).mockClear();
});

describe("default mint wiring", () => {
  it("getPersonalToken without deps drives the real capture driver, headless first", async () => {
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
    const token = await getPersonalToken({ store });
    expect(token).toBe("tok-minted");
    expect(capturePersonalSession).toHaveBeenCalledWith(
      expect.objectContaining({ headed: false, profileDir: personalProfileDir() }),
    );
    expect(store.read()?.accessToken).toBe("tok-minted");
  });

  it("the headed retry after a headless failure reaches the same driver", async () => {
    const mock = capturePersonalSession as ReturnType<typeof vi.fn>;
    mock
      .mockRejectedValueOnce(new Error("headless failed"))
      .mockResolvedValueOnce({
        accessToken: "tok-headed",
        chathubPath: "oid@tid",
        scenario: "S",
        variants: "v",
        licenseType: "Consumer",
        agentHost: "Bizchat.FullScreen",
        product: "Office",
        source: "officeweb",
        optionsSets: [],
        allowedMessageTypes: ["Chat"],
        expiresAt: Date.now() + 3_600_000,
        capturedAt: Date.now(),
      });
    const store = new PersonalStore(join(mkdtempSync(join(tmpdir(), "ps-")), "s.json"));
    store.write({
      accessToken: "tok-old",
      chathubPath: "oid@tid",
      scenario: "S",
      variants: "v",
      licenseType: "Consumer",
      agentHost: "Bizchat.FullScreen",
      product: "Office",
      source: "officeweb",
      optionsSets: [],
      allowedMessageTypes: ["Chat"],
      emailHint: null,
      probe: null,
      mintCount: 0,
      lastMintAt: 0,
      expiresAt: 1,
      capturedAt: 0,
    });
    const token = await getPersonalToken({ store });
    expect(token).toBe("tok-headed");
    expect(mock.mock.calls.map((call: unknown[]) => (call[0] as { headed: boolean }).headed)).toEqual([false, true]);
  });
});
