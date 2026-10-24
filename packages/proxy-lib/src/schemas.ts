import { z } from "zod/v4";
import { isPersonalMode, personalDefaultModel } from "@m365-copilot/core";

export const ToolCallFunction = z.object({
  name: z.string(),
  arguments: z.string(),
});

export const ToolCall = z.object({
  id: z.string(),
  type: z.literal("function").default("function"),
  function: ToolCallFunction,
});

export const ToolDefinition = z.object({
  type: z.literal("function").default("function"),
  function: z.object({
    name: z.string(),
    description: z.string().optional(),
    parameters: z.any().optional(),
  }),
});

export const ChatMessage = z.object({
  role: z.enum(["system", "developer", "user", "assistant", "tool"]).transform(
    (r) => (r === "developer" ? "system" : r),
  ),
  content: z.union([
    z.string(),
    z.array(
      z.object({
        type: z.string(),
        text: z.string().optional(),
        image_url: z.object({ url: z.string() }).passthrough().optional(),
      }),
    ),
  ]).nullable().optional(),
  tool_calls: z.array(ToolCall).optional(),
  tool_call_id: z.string().optional(),
  name: z.string().optional(),
});

export const ChatCompletionRequest = z.object({
  model: z.string().optional().default(() => (isPersonalMode() ? personalDefaultModel() : null) ?? "gpt-5.5-think-deeper"),
  messages: z.array(ChatMessage).min(1),
  stream: z.boolean().optional().default(false),
  stream_options: z.object({ include_usage: z.boolean().optional() }).optional(),
  temperature: z.number().optional(),
  max_tokens: z.number().optional(),
  tools: z.array(ToolDefinition).optional(),
  tool_choice: z.union([
    z.enum(["auto", "none", "required"]),
    z.object({
      type: z.literal("function"),
      function: z.object({ name: z.string() }),
    }),
  ]).optional(),
});
