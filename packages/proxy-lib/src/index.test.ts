import { describe, it, expect } from "vitest";
import { createApp } from "./index.js";

const LIVE = process.env.M365_LIVE === "1";

const tools = [
  {
    type: "function" as const,
    function: {
      name: "bash",
      description: "Run a shell command and return its output",
      parameters: {
        type: "object",
        properties: {
          command: { type: "string", description: "The command to run" },
        },
        required: ["command"],
      },
    },
  },
  {
    type: "function" as const,
    function: {
      name: "read_file",
      description: "Read the contents of a file at the given path",
      parameters: {
        type: "object",
        properties: {
          path: { type: "string", description: "Absolute path to the file" },
        },
        required: ["path"],
      },
    },
  },
];

function chatRequest(messages: Array<{ role: string; content?: string; tool_calls?: unknown[]; tool_call_id?: string; name?: string }>) {
  return new Request("http://localhost/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "m365-copilot",
      messages,
      tools,
    }),
  });
}

describe.skipIf(!LIVE)("proxy-lib e2e with tools (live)", () => {
  const app = createApp();

  it("should handle a chat completion with tools defined", async () => {
    const res = await app.fetch(
      chatRequest([
        { role: "user", content: "what npm deps do i have in this folder? use your bash and read_file tools to find out. start by running cat package.json" },
      ]),
    );

    expect(res.status).toBe(200);
    const body = await res.json();
    console.log("Turn 1:", JSON.stringify(body, null, 2));

    const choice = body.choices[0];
    expect(choice.message.role).toBe("assistant");

    if (choice.finish_reason === "tool_calls" && choice.message.tool_calls?.length > 0) {
      const toolCall = choice.message.tool_calls[0];
      console.log(`Tool call: ${toolCall.function.name}(${toolCall.function.arguments})`);

      let toolResult: string;
      if (toolCall.function.name === "bash") {
        const args = JSON.parse(toolCall.function.arguments);
        const { execSync } = await import("node:child_process");
        try {
          toolResult = execSync(args.command, { cwd: process.cwd(), encoding: "utf-8" }).trim();
        } catch (e: any) {
          toolResult = e.stderr || e.message;
        }
      } else if (toolCall.function.name === "read_file") {
        const args = JSON.parse(toolCall.function.arguments);
        const { readFileSync } = await import("node:fs");
        try {
          toolResult = readFileSync(args.path, "utf-8");
        } catch (e: any) {
          toolResult = e.message;
        }
      } else {
        toolResult = `Unknown tool: ${toolCall.function.name}`;
      }
      console.log("Tool result:", toolResult.slice(0, 500));

      const res2 = await app.fetch(
        chatRequest([
          { role: "user", content: "what npm deps do i have in this folder? use your bash and read_file tools to find out. start by running cat package.json" },
          {
            role: "assistant",
            content: choice.message.content,
            tool_calls: choice.message.tool_calls,
          },
          {
            role: "tool",
            tool_call_id: toolCall.id,
            name: toolCall.function.name,
            content: toolResult,
          },
        ]),
      );

      expect(res2.status).toBe(200);
      const body2 = await res2.json();
      console.log("Turn 2:", JSON.stringify(body2, null, 2));

      const choice2 = body2.choices[0];
      expect(choice2.message.content).toBeTruthy();
    } else {
      console.log("Model responded without tool calls (M365 Copilot overrode tool-calling protocol)");
      expect(choice.message.content).toBeTruthy();
    }
  }, 120_000);

  it("should return models list", async () => {
    const res = await app.fetch(new Request("http://localhost/v1/models"));
    expect(res.status).toBe(200);
    const body = await res.json();
    expect(body.object).toBe("list");
    expect(body.data.length).toBeGreaterThan(0);
    expect(body.data[0].id).toBeTruthy();
  });

  it("should return health check", async () => {
    const res = await app.fetch(new Request("http://localhost/health"));
    expect(res.status).toBe(200);
    const body = await res.json();
    expect(body.status).toBe("ok");
  });
});
