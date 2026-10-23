import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createMockRegistryAt } from "@m365-copilot/core";
import { routeFiles, startProxy, toRouteMethod, toRoutePath, type MountedProxy } from "./proxy-harness.js";
import { recordConsole } from "./setup.js";
import { ACCOUNTS_UI_HTML } from "../ui/accounts-page.js";

const ROOT = dirname(dirname(fileURLToPath(import.meta.url)));

const auth = vi.hoisted(() => ({ failure: null as string | null }));

vi.mock("@m365-copilot/core", async (importOriginal) => ({
  ...(await importOriginal<Record<string, unknown>>()),
  getToken: () => (auth.failure === null ? Promise.resolve("test-token") : Promise.reject(new Error(auth.failure))),
}));

const MODE_REFUSAL = "account management requires personal auth mode";

const REGISTRY_ACCOUNT_ROUTES = [
  "GET /accounts",
  "POST /accounts",
  "PATCH /accounts/:id",
  "POST /accounts/:id/activate",
  "DELETE /accounts/:id",
];

const DASHBOARD_ROUTE = "GET /accounts/ui";

const HEALTH_ROUTES = ["GET /health", "GET /healthz"];

const ACCOUNTS_FILES = [
  "routes/accounts/index.get.ts",
  "routes/accounts/index.post.ts",
  "routes/accounts/[id].patch.ts",
  "routes/accounts/[id].delete.ts",
  "routes/accounts/[id]/activate.post.ts",
  "routes/accounts/ui.get.ts",
];

const HEALTH_FILES = ["routes/health.get.ts", "routes/healthz.get.ts"];

const ENTERPRISE_CALLS: Array<[string, string, unknown]> = [
  ["GET", "/accounts", undefined],
  ["POST", "/accounts", { id: "home", label: "Home" }],
  ["PATCH", "/accounts/work", { label: "Renamed" }],
  ["POST", "/accounts/work/activate", undefined],
  ["DELETE", "/accounts/work", undefined],
  ["GET", "/accounts/ui", undefined],
];


const route = (file: string) => `${toRouteMethod(file)} ${toRoutePath(file)}`;

const samplePath = (file: string) => toRoutePath(file).replace(":id", "work");

interface Call {
  path: string;
  status: number;
  body: any;
  contentType: string | null;
}

function useRegistry(accounts: Parameters<typeof createMockRegistryAt>[1], activeId?: string): void {
  const dir = mkdtempSync(join(tmpdir(), "proxy-registry-"));
  process.env.M365_ACCOUNT_HOME = dir;
  const source = createMockRegistryAt(join(dir, "accounts.json"), accounts, activeId);
  process.env.M365_ACCOUNTS_FILE = source.file;
}

function readRegistry(): { activeId: string; accounts: Array<{ id: string; label: string }> } {
  return JSON.parse(readFileSync(process.env.M365_ACCOUNTS_FILE!, "utf8"));
}

let proxy: MountedProxy | null = null;

let bootFailure = false;

beforeEach(() => {
  delete process.env.M365_AUTH;
  useRegistry([
    { id: "default", hasSession: true, profileExists: true },
    { id: "work", cdpPort: 9444 },
  ]);
});

afterEach(async () => {
  await proxy?.close();
  proxy = null;
  delete process.env.M365_AUTH;
  delete process.env.M365_ACCOUNTS_FILE;
  delete process.env.M365_ACCOUNT_HOME;
});

async function bootProxy(mode: "personal" | "msal"): Promise<typeof call> {
  process.env.M365_AUTH = mode;
  proxy = await startProxy(ROOT);
  const tape = recordConsole();
  try {
    await proxy.boot();
  } catch {
    bootFailure = true;
  } finally {
    tape.restore();
  }
  return call;
}

async function bootWithConsole(mode: "personal" | "msal"): Promise<{ call: typeof call; logs: string[]; errors: string[] }> {
  process.env.M365_AUTH = mode;
  proxy = await startProxy(ROOT);
  const tape = recordConsole();
  bootFailure = false;
  try {
    await proxy.boot();
  } catch {
    bootFailure = true;
  } finally {
    tape.restore();
  }
  return { call, logs: tape.logs, errors: tape.errors };
}

async function bootDownProxy(mode: "msal"): Promise<typeof call> {
  auth.failure = "no credential";
  const booted = await bootWithConsole(mode);
  auth.failure = null;
  if (!bootFailure) throw new Error(`expected the auth plugin to abort boot in ${mode} mode`);
  bootFailure = false;
  return booted.call;
}

async function call(method: string, path: string, body?: unknown): Promise<Call> {
  const res = await proxy!.request(path, {
    method,
    headers: body === undefined ? {} : { "content-type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const text = await res.text();
  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }
  return { path: `${method} ${path}`, status: res.status, body: parsed, contentType: res.headers.get("content-type") };
}

describe("the harness mounts what Nitro routes", () => {
  it("discovers every account and health route file and maps each to its path and method", () => {
    expect(routeFiles(ROOT).filter((file) => ACCOUNTS_FILES.concat(HEALTH_FILES).includes(file)).map(route).sort()).toEqual(
      [...REGISTRY_ACCOUNT_ROUTES, DASHBOARD_ROUTE, ...HEALTH_ROUTES].sort(),
    );
  });

  it("binds the same method to each path Nitro generated a type for", () => {
    const generated = readFileSync(join(ROOT, ".nitro/types/nitro-routes.d.ts"), "utf8");
    for (const file of [...ACCOUNTS_FILES, ...HEALTH_FILES]) {
      const start = generated.indexOf(`'${toRoutePath(file)}': {`);
      const block = generated.slice(start, generated.indexOf("}", start));
      expect(block).toContain(`'${toRouteMethod(file).toLowerCase()}'`);
    }
  });
});

describe("GET /accounts/ui", () => {
  it("serves the dashboard as HTML in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("GET", "/accounts/ui");
    expect(r.status).toBe(200);
    expect(r.contentType).toBe("text/html; charset=utf-8");
    expect(r.body).toBe(ACCOUNTS_UI_HTML);
    expect(r.body).toContain("<!doctype html>");
    expect(r.body).toContain("<title>M365 accounts</title>");
    expect(r.body).toContain('id="rows"');
    expect(r.body).toContain("encodeURIComponent(id)");
  });

  it("is not found in enterprise mode", async () => {
    const enterprise = await bootProxy("msal");
    const r = await enterprise("GET", "/accounts/ui");
    expect(r.status).toBe(404);
    expect(r.body).toEqual({ error: { code: "not_found", message: MODE_REFUSAL } });
  });

  it("serves the file in ui/, and the package root no longer holds one", () => {
    expect(readFileSync(join(ROOT, "ui/accounts.html"), "utf8")).toBe(ACCOUNTS_UI_HTML);
    expect(existsSync(join(ROOT, "server-accounts-ui.ts"))).toBe(false);
    expect(readFileSync(join(ROOT, "routes/accounts/ui.get.ts"), "utf8")).toContain('from "../../ui/accounts-page"');
  });
});

describe("GET /accounts", () => {
  it("lists every account with its health and the registry path in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("GET", "/accounts");
    expect(r.status).toBe(200);
    expect(r.body.registry).toMatchObject({ path: process.env.M365_ACCOUNTS_FILE, activeId: "default", count: 2 });
    expect(r.body.accounts.map((a: any) => [a.id, a.active, a.profileExists])).toEqual([
      ["default", true, true],
      ["work", false, false],
    ]);
    expect(r.body.accounts[0].credential).toMatchObject({ status: expect.any(String) });
  });

  it("is not found in enterprise mode", async () => {
    const enterprise = await bootProxy("msal");
    const r = await enterprise("GET", "/accounts");
    expect(r.status).toBe(404);
    expect(r.body.error.code).toBe("not_found");
  });
});

describe("POST /accounts", () => {
  it("creates an account in personal mode and answers with both commands", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("POST", "/accounts", { id: "home", label: "Home (free)" });
    expect(r.status).toBe(201);
    expect(r.body.entry).toMatchObject({ id: "home", label: "Home (free)" });
    expect(r.body.chromeCommand).toContain(`--remote-debugging-port=${r.body.entry.cdpPort}`);
    expect(r.body.loginCommand).toBe("pnpm m365-personal login --account home");
    expect(readRegistry().accounts.map((a) => a.id)).toContain("home");
  });

  it("400s a bad id and 409s one that already exists in personal mode", async () => {
    const personal = await bootProxy("personal");
    const bad = await personal("POST", "/accounts", { id: "../x", label: "x" });
    expect(bad.status).toBe(400);
    expect(bad.body.error.code).toBe("invalid_id");
    const duplicate = await personal("POST", "/accounts", { id: "default", label: "x" });
    expect(duplicate.status).toBe(409);
    expect(duplicate.body.error.code).toBe("already_exists");
  });

  it("is not found in enterprise mode, and leaves the registry alone", async () => {
    const before = readRegistry();
    const enterprise = await bootProxy("msal");
    const r = await enterprise("POST", "/accounts", { id: "home", label: "Home" });
    expect(r.status).toBe(404);
    expect(r.body.error.code).toBe("not_found");
    expect(readRegistry()).toEqual(before);
  });
});

describe("PATCH /accounts/:id", () => {
  it("renames an account in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("PATCH", "/accounts/work", { label: "Work (free)" });
    expect(r.status).toBe(200);
    expect(r.body.entry).toMatchObject({ id: "work", label: "Work (free)" });
    expect(readRegistry().accounts.find((a) => a.id === "work")?.label).toBe("Work (free)");
  });

  it("400s an empty label in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("PATCH", "/accounts/work", { label: "" });
    expect(r.body.error.code).toBe("invalid_label");
  });

  it("is not found in enterprise mode, and renames nothing", async () => {
    const before = readRegistry();
    const enterprise = await bootDownProxy("msal");
    const r = await enterprise("PATCH", "/accounts/work", { label: "Renamed" });
    expect(r.status).toBe(404);
    expect(r.body.error.code).toBe("not_found");
    expect(readRegistry()).toEqual(before);
  });
});

describe("POST /accounts/:id/activate", () => {
  it("switches the active account in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("POST", "/accounts/work/activate");
    expect(r.status).toBe(200);
    expect(r.body).toEqual({ activeId: "work" });
    expect(readRegistry().activeId).toBe("work");
    const listed = await personal("GET", "/accounts");
    expect(listed.body.registry.activeId).toBe("work");
  });

  it("404s an unknown id in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("POST", "/accounts/gone/activate");
    expect(r.status).toBe(404);
    expect(r.body.error.code).toBe("not_found");
  });

  it("is not found in enterprise mode, and does not switch", async () => {
    const enterprise = await bootProxy("msal");
    const r = await enterprise("POST", "/accounts/work/activate");
    expect(r.status).toBe(404);
    expect(r.body.error.message).toBe(MODE_REFUSAL);
    expect(readRegistry().activeId).toBe("default");
  });
});

describe("DELETE /accounts/:id", () => {
  it("removes a quiet account in personal mode and reports what it forgot", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("DELETE", "/accounts/work");
    expect(r.status).toBe(200);
    expect(r.body).toEqual({ removed: "work", forgotten: 0 });
    expect(readRegistry().accounts.map((a) => a.id)).toEqual(["default"]);
  });

  it("409s the active account in personal mode", async () => {
    const personal = await bootProxy("personal");
    const r = await personal("DELETE", "/accounts/default");
    expect(r.status).toBe(409);
    expect(r.body.error.code).toBe("active");
  });

  it("is not found in enterprise mode, and removes nothing", async () => {
    const before = readRegistry();
    const enterprise = await bootProxy("msal");
    const r = await enterprise("DELETE", "/accounts/work");
    expect(r.status).toBe(404);
    expect(r.body.error.code).toBe("not_found");
    expect(readRegistry()).toEqual(before);
  });
});

describe("GET /health and GET /healthz", () => {
  it("answer in personal mode", async () => {
    const personal = await bootProxy("personal");
    const health = await personal("GET", "/health");
    expect(health.status).toBe(200);
    expect(health.body).toEqual({ status: "ok" });
    const healthz = await personal("GET", "/healthz");
    expect(healthz.status).toBe(200);
    expect(healthz.body).toMatchObject({
      status: "ok",
      authMode: "personal",
      accounts: { path: process.env.M365_ACCOUNTS_FILE, count: 2, activeId: "default", repository: "registry" },
    });
    expect(healthz.body.credential).toMatchObject({ status: expect.any(String) });
  });

  it("answer in enterprise mode, reporting the credential as not applicable", async () => {
    const enterprise = await bootProxy("msal");
    expect((await enterprise("GET", "/health")).body).toEqual({ status: "ok" });
    const healthz = await enterprise("GET", "/healthz");
    expect(healthz.status).toBe(200);
    expect(healthz.body).toMatchObject({ status: "ok", authMode: "msal", credential: { applicable: false } });
  });

  it("answer 200 with a report when the registry file is corrupt", async () => {
    writeFileSync(process.env.M365_ACCOUNTS_FILE!, "{ not json");
    const personal = await bootProxy("personal");
    const r = await personal("GET", "/healthz");
    expect(r.status).toBe(200);
    expect(r.body.accounts.problem).toMatchObject({ code: expect.any(String), hint: expect.any(String) });
  });
});

describe("the startup auth plugin", () => {
  afterEach(() => {
    auth.failure = null;
  });

  it("reports a failed login in personal mode and boots the server anyway", async () => {
    auth.failure = "no credential";
    const { call: personal, logs, errors } = await bootWithConsole("personal");
    expect(logs).toEqual(["M365_AUTH=personal"]);
    expect(errors[0]).toContain("Auth failed:");
    expect(errors.join("\n")).toContain("Booting anyway");
    expect((await personal("GET", "/healthz")).body.authMode).toBe("personal");
  });

  it("keeps the server off the account routes in enterprise mode when the login fails", async () => {
    auth.failure = "no credential";
    const { call: enterprise, errors } = await bootWithConsole("msal");
    expect(bootFailure).toBe(true);
    expect(errors[0]).toContain("Auth failed:");
    expect(errors.join("\n")).not.toContain("Booting anyway");
    const answered: Array<{ route: string; status: number; code: string }> = [];
    for (const [method, path, body] of ENTERPRISE_CALLS) {
      const r = await enterprise(method, path, body);
      answered.push({ route: r.path, status: r.status, code: r.body.error?.code });
    }
    expect(answered).toEqual(
      ENTERPRISE_CALLS.map(([method, path]) => ({ route: `${method} ${path}`, status: 404, code: "not_found" })),
    );
  });
});

describe("the account surface in enterprise mode", () => {
  it("answers not-found to every account call, dashboard included", async () => {
    const enterprise = await bootProxy("msal");
    const answered: Array<{ route: string; status: number; code: string }> = [];
    for (const [method, path, body] of ENTERPRISE_CALLS) {
      const r = await enterprise(method, path, body);
      answered.push({ route: r.path, status: r.status, code: r.body.error?.code });
    }
    expect(answered).toEqual(
      ENTERPRISE_CALLS.map(([method, path]) => ({ route: `${method} ${path}`, status: 404, code: "not_found" })),
    );
  });
});
