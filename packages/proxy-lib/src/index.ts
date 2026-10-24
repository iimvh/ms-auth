import { type ModelSessionOptions, getAvailableModels, isPersonalMode, ensurePersonalCredentials, resolveAccount } from "@m365-copilot/core";
import { ChatCompletionRequest } from "./schemas.js";
import { SessionPool, handleChatCompletion } from "./handler.js";

export { SessionPool, handleChatCompletion } from "./handler.js";
export { buildHealthzPayload, type HealthzPayload } from "./healthz.js";
export {
  accountsActivateResponse,
  accountsAddResponse,
  accountsDeleteResponse,
  accountsListResponse,
  accountsRenameResponse,
  listAccounts,
  resetAccountsHealthCache,
  toRouteResult,
  type AccountHealthView,
  type AccountsDeps,
  type AccountsListBody,
  type RouteResult,
} from "./accounts-api.js";
export { ChatCompletionRequest, ChatMessage, ToolCall, ToolDefinition } from "./schemas.js";

export {
  formatMessages,
  formatToolDefinitions,
  parseToolCalls,
  getMessageContent,
  type Message,
  type ToolDef,
  type ToolChoice,
  type ParsedToolCall,
  type ParseResult,
} from "@m365-copilot/core";

export const HEALTH_PAYLOAD = { status: "ok" } as const;

const CONTEXT_WINDOW_TOKENS = Number(process.env.M365_CONTEXT_WINDOW) || 1_000_000;
const MAX_OUTPUT_TOKENS = Number(process.env.M365_MAX_OUTPUT_TOKENS) || 1_000_000;

export function buildModelsPayload() {
  const created = Math.floor(Date.now() / 1000);
  return {
    object: "list",
    data: getAvailableModels().map((id) => ({
      id,
      object: "model",
      created,
      owned_by: "microsoft",
      context_window: CONTEXT_WINDOW_TOKENS,
      max_context_length: CONTEXT_WINDOW_TOKENS,
      max_input_tokens: CONTEXT_WINDOW_TOKENS,
      max_output_tokens: MAX_OUTPUT_TOKENS,
    })),
  };
}

const CORS_HEADERS: Record<string, string> = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
};

function withCors(res: Response): Response {
  for (const [k, v] of Object.entries(CORS_HEADERS)) res.headers.set(k, v);
  return res;
}

function json(status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

export interface FetchApp {
  fetch(req: Request): Promise<Response>;
}

export interface CreateAppOptions extends Omit<ModelSessionOptions, "accountId"> {
  accountId?: () => string | null;
}

function activeAccountId(): string | null {
  return isPersonalMode() ? resolveAccount().id : null;
}

export function createApp(options: CreateAppOptions = {}): FetchApp {
  const { accountId: accountIdOf = activeAccountId, ...sessionOptions } = options;
  const pool = new SessionPool(sessionOptions);

  async function fetch(req: Request): Promise<Response> {
    const { pathname } = new URL(req.url);
    const method = req.method.toUpperCase();

    if (method === "OPTIONS") {
      return withCors(new Response(null, { status: 204 }));
    }

    if (method === "GET" && pathname === "/health") {
      return withCors(json(200, HEALTH_PAYLOAD));
    }

    if (method === "GET" && pathname === "/v1/models") {
      return withCors(json(200, buildModelsPayload()));
    }

    if (method === "POST" && pathname === "/v1/chat/completions") {
      let body: ReturnType<typeof ChatCompletionRequest.parse>;
      try {
        body = ChatCompletionRequest.parse(await req.json());
      } catch (err: any) {
        return withCors(
          json(400, { error: { message: err.message, type: "invalid_request_error" } }),
        );
      }
      const personal = isPersonalMode();
      return withCors(await handleChatCompletion(body, pool, {
        signal: req.signal,
        accountId: accountIdOf(),
        ensureCredentials: personal ? (id) => ensurePersonalCredentials(id ?? undefined) : undefined,
      }));
    }

    return withCors(
      json(404, { error: { message: "Not found", type: "invalid_request_error" } }),
    );
  }

  return { fetch };
}
