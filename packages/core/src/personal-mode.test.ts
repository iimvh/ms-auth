import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { isPersonalMode, personalDefaultModel, personalModeDefaults, resolveAuthMode, shouldSendCodeInterpreter } from "./personal-mode.js";
import { PersonalStore } from "./personal-session.js";
import { accountsDir, activateAccount, clearMockRegistry, useMockRegistry } from "./account-registry.js";
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

beforeEach(() => {
  // personalDefaultModel's record path goes through the env-resolved session
  // file; point it at a scratch file so test isolation holds.
  process.env.M365_PERSONAL_SESSION_FILE = join(mkdtempSync(join(tmpdir(), "ps-")), "s.json");
});

afterEach(() => {
  delete process.env.M365_AUTH;
  delete process.env.M365_PERSONAL_CODE_INTERPRETER;
  delete process.env.M365_PERSONAL_MODEL;
  delete process.env.M365_PERSONAL_SESSION_FILE;
});

describe("auth mode", () => {
  it("defaults to msal and accepts personal case-insensitively", () => {
    expect(resolveAuthMode()).toBe("msal");
    process.env.M365_AUTH = "Personal";
    expect(isPersonalMode()).toBe(true);
  });
  it("fails loudly on an unknown value", () => {
    process.env.M365_AUTH = "microsoft";
    expect(() => resolveAuthMode()).toThrow(/msal.*personal|personal.*msal/);
  });
});

describe("personal defaults", () => {
  it("is agent-less, sandbox-free and image-free", () => {
    process.env.M365_AUTH = "personal";
    expect(personalModeDefaults()).toEqual({ useAgent: false, codeInterpreter: false, imageGeneration: false });
  });
  it("re-enables the sandbox only on request", () => {
    process.env.M365_AUTH = "personal";
    process.env.M365_PERSONAL_CODE_INTERPRETER = "1";
    expect(personalModeDefaults().codeInterpreter).toBe(true);
  });
  it("shouldSendCodeInterpreter is false only for personal+agent-less", () => {
    expect(shouldSendCodeInterpreter(null, false)).toBe(true);
    expect(shouldSendCodeInterpreter(null, true)).toBe(false);
    expect(shouldSendCodeInterpreter("agent-1", true)).toBe(false);
    process.env.M365_PERSONAL_CODE_INTERPRETER = "1";
    expect(shouldSendCodeInterpreter(null, true)).toBe(true);
  });
  it("personalDefaultModel prefers the env override, then the probed model", () => {
    expect(personalDefaultModel()).toBeNull();
    expect(personalDefaultModel({ model: "claude-sonnet", ok: true, at: 1 })).toBe("claude-sonnet");
    process.env.M365_PERSONAL_MODEL = "gpt-5.5";
    expect(personalDefaultModel({ model: "claude-sonnet", ok: true, at: 1 })).toBe("gpt-5.5");
  });

  it("personalDefaultModel with no argument reads the record's probe", () => {
    const store = new PersonalStore(process.env.M365_PERSONAL_SESSION_FILE!);
    expect(personalDefaultModel()).toBeNull(); // no record
    const probe = { model: "claude-sonnet", ok: true, at: 1 };
    store.write({
      accessToken: "t", chathubPath: "p", scenario: "S", variants: "v", licenseType: "Consumer",
      agentHost: "h", product: "Office", source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
      emailHint: null, probe, mintCount: 1, lastMintAt: 0, expiresAt: 1, capturedAt: 0,
    });
    expect(personalDefaultModel()).toBe("claude-sonnet");
    const bad = { model: "claude-sonnet", ok: false, at: 1 };
    store.write({
      accessToken: "t", chathubPath: "p", scenario: "S", variants: "v", licenseType: "Consumer",
      agentHost: "h", product: "Office", source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
      emailHint: null, probe: bad, mintCount: 1, lastMintAt: 0, expiresAt: 1, capturedAt: 0,
    });
    expect(personalDefaultModel()).toBeNull();
  });
});

describe("personalDefaultModel follows the active account", () => {
  it("reads the ACTIVE account's probe after a switch", () => {
    delete process.env.M365_PERSONAL_MODEL;
    delete process.env.M365_PERSONAL_SESSION_FILE; // beforeAll points it at a scratch file; the registry must win
    useMockRegistry([{ id: "default", hasSession: true }, { id: "work", hasSession: true }]);
    const rec = (model: string) => ({ accessToken: "t", chathubPath: "o@t", scenario: "S", variants: "v",
      licenseType: "Free", agentHost: "h", product: null, source: "officeweb", optionsSets: [],
      allowedMessageTypes: ["Chat"], emailHint: null, probe: { model, ok: true, at: 1 },
      mintCount: 1, lastMintAt: 0, expiresAt: 1, capturedAt: 0 });
    new PersonalStore(join(accountsDir(), "default", "session.json")).write(rec("claude-sonnet"));
    new PersonalStore(join(accountsDir(), "work", "session.json")).write(rec("gpt-5.5"));
    expect(personalDefaultModel()).toBe("claude-sonnet");
    activateAccount("work");
    expect(personalDefaultModel()).toBe("gpt-5.5");
    clearMockRegistry();
  });

  it("M365_PERSONAL_MODEL still wins over the active account's probe", () => {
    useMockRegistry([{ id: "work", hasSession: true }]);
    new PersonalStore(join(accountsDir(), "work", "session.json")).write({
      accessToken: "t", chathubPath: "o@t", scenario: "S", variants: "v", licenseType: "Free",
      agentHost: "h", product: null, source: "officeweb", optionsSets: [], allowedMessageTypes: ["Chat"],
      emailHint: null, probe: { model: "gpt-5.5", ok: true, at: 1 }, mintCount: 1, lastMintAt: 0,
      expiresAt: 1, capturedAt: 0 });
    process.env.M365_PERSONAL_MODEL = "claude-sonnet";
    delete process.env.M365_PERSONAL_SESSION_FILE;
    expect(personalDefaultModel()).toBe("claude-sonnet");
    clearMockRegistry();
  });
});