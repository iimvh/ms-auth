import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  accountsActivateResponse,
  accountsAddResponse,
  accountsDeleteResponse,
  accountsListResponse,
  accountsRenameResponse,
  resetAccountsHealthCache,
  type AccountsDeps,
} from "./accounts-api.js";
import { clearMockRegistry, useMockRegistry } from "@m365-copilot/core";

const deps = (live: string[] = []): AccountsDeps => ({
  pool: {
    liveAccountIds: () => live,
    forgetAccount: () => 1,
  },
  now: () => 1_000_000_000,
  probeCdp: async () => false,
});

beforeEach(() => {
  // useMockRegistry writes under M365_ACCOUNT_HOME's accountsDir(); sharing a
  // home with another test file would make this registry collide with its own.
  process.env.M365_ACCOUNT_HOME = mkdtempSync(join(tmpdir(), "accounts-api-"));
  clearMockRegistry();
  resetAccountsHealthCache();
  delete process.env.M365_AUTH;
  process.env.M365_AUTH = "personal";
});

describe("GET /accounts", () => {
  it("lists every account with health, profile presence and the registry path", async () => {
    useMockRegistry([{ id: "default", hasSession: true, profileExists: true }, { id: "work" }]);
    const r = await accountsListResponse(deps());
    expect(r.status).toBe(200);
    const body = r.body as { registry: { path: string }; accounts: any[] };
    expect(body.registry.path).toBe(process.env.M365_ACCOUNTS_FILE);
    expect(body.accounts).toHaveLength(2);
    expect(body.accounts[0]).toMatchObject({ id: "default", active: true, profileExists: true });
    expect(body.accounts[0].credential.status).toBeTruthy();
    expect(body.accounts[1]).toMatchObject({ id: "work", active: false });
  });

  it("caches CDP probes so a dashboard refresh opens no new sockets", async () => {
    useMockRegistry([{ id: "default" }]);
    const probe = vi.fn(async () => false);
    await accountsListResponse({ ...deps(), probeCdp: probe });
    await accountsListResponse({ ...deps(), probeCdp: probe });
    expect(probe).toHaveBeenCalledTimes(1);
  });
});

describe("POST /accounts", () => {
  it("201 with the chrome command and login command", async () => {
    useMockRegistry([{ id: "default" }]);
    const r = await accountsAddResponse({ id: "work", label: "Work" }, deps());
    expect(r.status).toBe(201);
    const body = r.body as any;
    expect(body.chromeCommand).toContain(`--remote-debugging-port=${body.entry.cdpPort}`);
    expect(body.loginCommand).toBe("pnpm m365-personal login --account work");
  });

  it("400 for a bad id, 409 for one that exists", async () => {
    useMockRegistry([{ id: "default" }]);
    expect((await accountsAddResponse({ id: "../x", label: "x" }, deps())).status).toBe(400);
    expect((await accountsAddResponse({ id: "default", label: "x" }, deps())).status).toBe(409);
  });

  it("400 for an empty label", async () => {
    useMockRegistry([{ id: "default" }]);
    expect((await accountsAddResponse({ id: "work", label: "  " }, deps())).status).toBe(400);
  });
});

describe("DELETE /accounts/:id", () => {
  it("refuses the active account and the last one", async () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    await accountsActivateResponse("work", deps());
    expect(accountsDeleteResponse("work", deps()).result.status).toBe(409);
    useMockRegistry([{ id: "only" }]);
    expect(accountsDeleteResponse("only", deps()).result.status).toBe(409);
  });

  it("refuses while a conversation of that account is live, naming the fix", async () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    const r = accountsDeleteResponse("work", deps(["work"])).result;
    expect(r.status).toBe(409);
    const body = r.body as any;
    expect(body.error.code).toBe("account_busy");
    expect(body.error.message).toContain('"work"');
    // The hint must name a way out that exists: the page cannot finish a
    // thread, so it names the idle limit and switching the client.
    expect(body.error.hint).toContain("M365_SESSION_IDLE_MS");
    expect(body.error.hint).toContain("switch your client");
  });

  it("removes a quiet account and forgets its leftover pool state", async () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    const { result, forgotten } = accountsDeleteResponse("work", deps());
    expect(result.status).toBe(200);
    expect(forgotten).toBe(1);
  });
});

describe("PATCH and activate", () => {
  it("renames and activates; both 404 an unknown id", async () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    expect(accountsRenameResponse("work", { label: "Work (free)" }, deps()).status).toBe(200);
    expect(accountsActivateResponse("work", deps()).status).toBe(200);
    expect(accountsRenameResponse("gone", { label: "x" }, deps()).status).toBe(404);
    expect(accountsActivateResponse("gone", deps()).status).toBe(404);
  });

  it("400s an empty label on rename", async () => {
    useMockRegistry([{ id: "default" }, { id: "work" }]);
    expect(accountsRenameResponse("work", { label: "" }, deps()).status).toBe(400);
  });
});

describe("MSAL mode", () => {
  it("every route is 404", async () => {
    delete process.env.M365_AUTH;
    useMockRegistry([{ id: "default" }]);
    expect((await accountsListResponse(deps())).status).toBe(404);
    expect((await accountsAddResponse({ id: "work", label: "x" }, deps())).status).toBe(404);
    expect(accountsRenameResponse("default", { label: "x" }, deps()).status).toBe(404);
    expect(accountsActivateResponse("default", deps()).status).toBe(404);
    expect(accountsDeleteResponse("default", deps()).result.status).toBe(404);
  });
});
