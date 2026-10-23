import { describe, expect, it } from "vitest";
import { SessionPool } from "./handler.js";

const msgs = [{ role: "user", content: "open the file" }] as never;

describe("SessionPool and the account", () => {
  it("the same opener under two accounts is two conversations", () => {
    const pool = new SessionPool();
    const a = pool.resolve(msgs, undefined, "default");
    const b = pool.resolve(msgs, undefined, "work");
    expect(a).not.toBe(b);
    expect(a.accountId).toBe("default");
    expect(b.accountId).toBe("work");
    expect(pool.size).toBe(2);
  });

  it("a continuation under the SAME account still resolves to its conversation", () => {
    const pool = new SessionPool();
    const a = pool.resolve(msgs, undefined, "work");
    expect(pool.resolve([...msgs], undefined, "work")).toBe(a);
  });

  it("the builtin layout (no account) keys as before", () => {
    const pool = new SessionPool();
    const a = pool.resolve(msgs, undefined, null);
    expect(pool.resolve(msgs, undefined, undefined)).toBe(a);
    expect(a.accountId).toBeNull();
  });

  it("liveAccountIds and forgetAccount", () => {
    const pool = new SessionPool();
    pool.resolve(msgs, undefined, "a");
    pool.resolve([{ role: "user", content: "other" }] as never, undefined, "b");
    expect(pool.liveAccountIds().sort()).toEqual(["a", "b"]);
    expect(pool.forgetAccount("a")).toBe(1);
    expect(pool.liveAccountIds()).toEqual(["b"]);
  });
});
