import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

process.env.M365_SECRETS_FILE = join(mkdtempSync(join(tmpdir(), "sec-")), "secrets.json");

const { SecretsFileSchema, loadSecrets } = await import("./auth.js");

const SECRETS_FILE = process.env.M365_SECRETS_FILE;

describe("the secrets file", () => {
  it("is a shape: email, password and mfaSecret, each a non-empty string", () => {
    expect(SecretsFileSchema.safeParse({ email: "a@b.c", password: "p", mfaSecret: "JBSWY3DPEHPK3PXP" }).success).toBe(true);
    expect(SecretsFileSchema.safeParse({ email: "", password: "p", mfaSecret: "s" }).success).toBe(false);
    expect(SecretsFileSchema.safeParse({ email: "a@b.c", password: "", mfaSecret: "s" }).success).toBe(false);
    expect(SecretsFileSchema.safeParse({ email: "a@b.c", password: "p", mfaSecret: "" }).success).toBe(false);
    expect(SecretsFileSchema.safeParse({ email: "a@b.c", password: "p" }).success).toBe(false);
    expect(SecretsFileSchema.safeParse({ email: 1, password: "p", mfaSecret: "s" }).success).toBe(false);
    expect(SecretsFileSchema.safeParse(null).success).toBe(false);
  });

  it("loads a valid file, and loads nothing from an invalid one", () => {
    writeFileSync(SECRETS_FILE, JSON.stringify({ email: "a@b.c", password: "p", mfaSecret: "JBSWY3DPEHPK3PXP" }));
    expect(loadSecrets()).toEqual({ email: "a@b.c", password: "p", mfaSecret: "JBSWY3DPEHPK3PXP" });

    writeFileSync(SECRETS_FILE, JSON.stringify({ email: "a@b.c", password: "p" }));
    expect(loadSecrets()).toBeNull();

    writeFileSync(SECRETS_FILE, JSON.stringify({ email: "a@b.c", password: "p", mfaSecret: 0 }));
    expect(loadSecrets()).toBeNull();

    const { email } = { email: "a@b.c" };
    writeFileSync(SECRETS_FILE, JSON.stringify({ [email]: "a@b.c", password: "p", mfaSecret: 0 }));
    expect(loadSecrets()).toBeNull();

    writeFileSync(SECRETS_FILE, "{not json");
    expect(loadSecrets()).toBeNull();
  });
});
