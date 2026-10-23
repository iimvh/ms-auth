import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { buildHealthzPayload } from "./healthz.js";
import { activateAccount, clearMockRegistry, useMockRegistry } from "@m365-copilot/core";

beforeEach(() => {
  process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "healthz-"));
});

afterEach(() => {
  delete process.env.M365_AUTH;
  clearMockRegistry();
  delete process.env.M365_ACCOUNT_HOME;
});

describe("/healthz", () => {
  it("reports only the ACTIVE account plus counts, never every CDP", async () => {
    process.env.M365_AUTH = "personal";
    useMockRegistry([{ id: "default", hasSession: true, cdpPort: 9333 }, { id: "work", cdpPort: 9222 }]);
    const payload = await buildHealthzPayload({ probeCdp: async () => false });
    expect(payload.accounts).toMatchObject({ count: 2, activeId: "default" });
    expect(payload.credential).toMatchObject({ status: expect.any(String) });
  });

  it("follows a switch without a restart", async () => {
    process.env.M365_AUTH = "personal";
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    activateAccount("work");
    expect((await buildHealthzPayload({ probeCdp: async () => false })).accounts.activeId).toBe("work");
  });

  it("is MSAL-shaped in enterprise mode (plus the registry counts)", async () => {
    delete process.env.M365_AUTH;
    useMockRegistry([{ id: "only" }]);
    const payload = await buildHealthzPayload();
    expect(payload).toMatchObject({ status: "ok", authMode: "msal", credential: { applicable: false } });
    expect(payload.accounts.activeId).toBeTruthy(); // the registry, when there is one
  });
});
