import { describe, expect, it, vi } from "vitest";
import { handleChatCompletion, SessionPool } from "./handler.js";

const body = (stream: boolean) => ({ model: "claude-sonnet", stream,
  messages: [{ role: "user", content: "hi" }] }) as never;
const dead = async () => { throw new Error("M365 personal session expired — run: pnpm m365-personal login"); };

describe("pre-flight credentials", () => {
  it.each([true, false])("stream=%s gets a real 401 JSON before any SSE byte", async (stream) => {
    const res = await handleChatCompletion(body(stream), new SessionPool(), { ensureCredentials: dead });
    expect(res.status).toBe(401);
    expect(res.headers.get("content-type")).toContain("application/json");
    const payload = await res.json();
    expect(payload.error.code).toBe("m365_personal_session_dead");
    expect(payload.error.message).toContain("pnpm m365-personal login");
  });

  it("a successful pre-flight does not change behaviour", async () => {
    const ensure = vi.fn(async () => {});
    const res = await handleChatCompletion(body(true), new SessionPool({ getToken: async () => "tok" }), { ensureCredentials: ensure });
    expect(ensure).toHaveBeenCalledTimes(1);
    expect(res.status).toBe(200);
  });
});
