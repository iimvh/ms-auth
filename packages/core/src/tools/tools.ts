import { createLogger } from "../io/log.js";
import {
  buildSpecMap,
  currentFramingVariant,
  deriveFencedSpec,
  findFirstToolFence,
  formatFencedToolDefinitions,
  parseFencedToolCalls,
  renderFencedCall,
  transcriptStyleForVariant,
} from "../framing/fenced.js";

const log = createLogger("tools");

export interface ToolFunction {
  name: string;
  description?: string;
  parameters?: {
    properties?: Record<string, { type?: string; [k: string]: unknown }>;
    required?: string[];
    [k: string]: unknown;
  };
}

export interface ToolDef {
  type?: string;
  function: ToolFunction;
}

export interface Message {
  role: string;
  content?: string | Array<{ type: string; text?: string; image_url?: { url: string } }> | null;
  tool_calls?: Array<{
    id: string;
    function: { name: string; arguments: string };
  }>;
  tool_call_id?: string;
  name?: string;
}

export type ToolChoice =
  | "auto"
  | "none"
  | "required"
  | { type: "function"; function: { name: string } }
  | undefined;

const TOOL_CALL_REGEX = /\{\s*"tool"\s*:\s*"[^"]+"\s*,\s*"arguments"\s*:\s*\{[\s\S]*?\}\s*\}/g;
const FENCED_TOOL_CALL_REGEX = /```tool_call\s*\n(\{[\s\S]*?\})\s*\n\s*```/g;

const CONFIDENCE_REGEX = /\{\s*"confidence"\s*:\s*-?[0-9.]+\s*\}/g;
const FINAL_OBJECT_REGEX = /\{\s*"final"\s*:\s*"(?:[^"\\]|\\.)*"\s*\}/g;

function cleanLooseText(text: string): string | null {
  let out = text;
  for (const m of out.match(FINAL_OBJECT_REGEX) ?? []) {
    try {
      const value = JSON.parse(m).final;
      if (typeof value === "string") out = out.replace(m, value);
    } catch {
    }
  }
  out = out.replace(CONFIDENCE_REGEX, "").trim();
  return out.length ? out : null;
}

export function formatToolDefinitions(tools: ToolDef[], variantOverride?: string): string {
  return formatFencedToolDefinitions(tools, variantOverride);
}

export function formatToolChoiceInstruction(toolChoice: ToolChoice): string {
  if (!toolChoice || toolChoice === "auto") return "";
  if (toolChoice === "none") return "\nDo NOT call tools. Text only.";
  if (toolChoice === "required") return "\nYou MUST call at least one tool.";
  if (typeof toolChoice === "object" && toolChoice.function) {
    return `\nYou MUST call "${toolChoice.function.name}".`;
  }
  return "";
}

export function getMessageContent(msg: Message): string {
  if (msg.content === null || msg.content === undefined) return "";
  if (typeof msg.content === "string") return msg.content;
  return msg.content.map((p) => p.text || "").join("");
}

function toolCallSummary(rawArgs: string): string {
  let args: Record<string, unknown> = {};
  try {
    args = typeof rawArgs === "string" ? JSON.parse(rawArgs || "{}") : (rawArgs ?? {});
  } catch {
    return "";
  }
  const primary =
    args.command ?? args.cmd ?? args.script ?? args.path ?? args.file ??
    args.filename ?? args.query ?? Object.values(args).find((v) => typeof v === "string");
  if (typeof primary !== "string") return "";
  return primary.replace(/\s+/g, " ").replace(/"/g, "'").trim().slice(0, 100);
}

export function formatToolResponse(m: Message, history: Message[]): string {
  let meta: { name: string; summary: string } | undefined;
  if (m.tool_call_id) {
    for (const a of history) {
      const tc = a.role === "assistant" ? a.tool_calls?.find((c) => c.id === m.tool_call_id) : undefined;
      if (tc) {
        meta = { name: tc.function.name, summary: toolCallSummary(tc.function.arguments) };
        break;
      }
    }
  }
  const name = m.name || meta?.name || "tool";
  const cmdAttr = meta?.summary ? ` command="${meta.summary}"` : "";
  return `<tool_response tool="${name}"${cmdAttr}>\n${getMessageContent(m)}\n</tool_response>`;
}

function maybeInjectReplyTool(tools: ToolDef[]): ToolDef[] {
  const enabled = process.env.M365_INJECT_REPLY_TOOL || currentFramingVariant() === "reply_tool";
  if (!enabled) return tools;
  if (tools.some((t) => t.function.name === "reply")) return tools;
  const replyTool: ToolDef = {
    type: "function",
    function: {
      name: "reply",
      description:
        "Send a plain-text answer to the user. Use this whenever you would otherwise reply in prose.",
      parameters: {
        type: "object",
        properties: { text: { type: "string", description: "The text to send" } },
        required: ["text"],
      },
    },
  };
  return [replyTool, ...tools];
}

export function formatMessages(
  messages: Message[],
  tools?: ToolDef[],
  toolChoice?: ToolChoice,
  conversationId?: string,
  framingVariant?: string,
): string {
  const parts: string[] = [];

  if (conversationId) {
    parts.push(`<conversation_id>${conversationId}</conversation_id>`);
  }

  const effectiveTools = tools ? maybeInjectReplyTool(tools) : tools;
  const specMap = effectiveTools ? buildSpecMap(effectiveTools) : null;
  const style = transcriptStyleForVariant(framingVariant ?? currentFramingVariant());
  if (effectiveTools && effectiveTools.length > 0 && toolChoice !== "none") {
    const framing = `${formatToolDefinitions(effectiveTools, framingVariant)}${formatToolChoiceInstruction(toolChoice)}`;
    parts.push(style.framingTag ? `<${style.framingTag}>\n${framing}\n</${style.framingTag}>` : framing);
  }

  for (const m of messages) {
    if (m.role === "assistant" && m.tool_calls && m.tool_calls.length > 0) {
      const calls = m.tool_calls.map((tc) => {
        const rawArgs = tc.function.arguments;
        let argsObj: Record<string, unknown> = {};
        try {
          argsObj = typeof rawArgs === "string" ? JSON.parse(rawArgs || "{}") : (rawArgs ?? {});
        } catch {
        }
        const spec = specMap?.get(tc.function.name) ?? deriveFencedSpec({
          type: "function",
          function: {
            name: tc.function.name,
            parameters: {
              properties: Object.fromEntries(
                Object.keys(argsObj).map((k) => [k, { type: "string" }]),
              ),
            },
          },
        });
        return renderFencedCall(spec, argsObj);
      }).join("\n");
      const content = getMessageContent(m);
      parts.push(`<assistant>${content ? "\n" + content : ""}\n${calls}\n</assistant>`);
    } else if (m.role === "tool") {
      parts.push(formatToolResponse(m, messages));
    } else if (m.role === "system") {
      parts.push(`<${style.systemTag}>\n${getMessageContent(m)}\n</${style.systemTag}>`);
    } else {
      parts.push(`<${m.role}>\n${getMessageContent(m)}\n</${m.role}>`);
    }
  }

  return parts.join("\n\n");
}

export interface ParsedToolCall {
  id: string;
  type: "function";
  function: { name: string; arguments: string };
}

export interface ParseResult {
  hasToolCalls: boolean;
  toolCalls: ParsedToolCall[];
  textContent: string | null;
}

const CONFABULATION_PATTERNS: RegExp[] = [
  /return(?:ing|s|ed)?\s+no\s+(?:output|results?|content)/i,
  /no\s+(?:output|results?|content|data)\s+(?:was\s+|were\s+)?(?:return|provid|present)/i,
  /(?:unable|not able|can.?t|cannot)\s+(?:to\s+)?(?:access|inspect|list|read|run|execute|retrieve|fetch|locate|see|open)/i,
  /don.?t\s+have\s+access/i,
  /no\s+(?:longer\s+have|access\s+to)/i,
  /lost\s+(?:access|my\s+access|the\s+ability)/i,
  /(?:restart|start\s+over|begin\s+again|re-?run)\s+(?:the\s+|this\s+)?(?:task|session|conversation|work)\s+in\s+(?:a\s+)?/i,
  /(?:in|use|switch\s+to|need)\s+(?:a\s+)?(?:different|another|proper|coding-?enabled|tool-?enabled|shell-?enabled)\s+(?:session|environment|conversation|mode)/i,
  /(?:tool|editor|shell|command|file-?editing)s?[^.\n]{0,40}\b(?:not\s+available|unavailable|aren.?t\s+available|isn.?t\s+available|are\s+not\s+accessible)/i,
  /(?:can.?t|cannot|not\s+able\s+to|unable\s+to)\s+(?:directly\s+)?(?:edit|modify|write\s+to|change|save|create|open)\s+(?:the\s+|any\s+|to\s+)?files?/i,
  /paste\s+(?:the\s+)?(?:contents?|files?|code|them)/i,
  /provide\s+(?:the\s+)?(?:contents?|files?)/i,
  /(?:environment|shell|tool)\s+(?:isn.?t|is not|aren.?t|are not|appears? to be)\s+(?:return|provid|respond|work|access)/i,
  /(?:current\s+working\s+directory|working\s+directory|cwd)[\s\S]{0,120}\/mnt\/data/i,
  /(?:\bpwd\b|\bcd\b)[\s\S]{0,60}\/mnt\/data/i,
  /(?:ran|used|executed|called)[\s\S]{0,80}container\.(?:exec|open_image|download)/i,
  /container\.(?:exec|open_image|download)[\s\S]{0,120}(?:returned|output|shows?|result)/i,
  /no\s+files?\s+(?:in|found|present|visible)/i,
  /(?:file|directory|folder|it)\s+(?:appears?|seems?|looks?)\s+(?:to\s+be\s+)?empty/i,
  /nothing\s+to\s+(?:simplify|fix|do|change|show|read)/i,
  /(?:tool|command|it)\s+returned\s+(?:no|empty|nothing)/i,
  /(?:session|environment|runtime)\s+(?:does\s+not|doesn.?t|cannot)\s+(?:expose|mount|provide)\s+(?:the\s+)?(?:local\s+)?(?:repository\s+)?filesystem/i,
  /(?:my|the)\s+filesystem\s+(?:only\s+)?(?:contained|contains|has)[\s\S]{0,80}\/mnt\/data/i,
  /bu\s+oturumda(?:[^.\n]|\.(?=\S)){0,80}(?:devre\s*dışı|etkin\s+değil|kullanılamıyor|mevcut\s+değil)/i,
  /(?:özelliğim|yeteneğim|aracım|araçlarım|erişimim|yetkim|iznim)(?:[^.\n]|\.(?=\S)){0,30}(?:devre\s*dışı|etkin\s+değil|yok|bulunmuyor)/i,
  /(?:oluştur|üret|hazırla|dönüştür|tamamla|kur|güncelle|kayded|eriş|çalıştır|oku|yaz|düzenle|incele|listele|aç)y?[ae]m(?:ıyorum|iyorum|am|em|adım|edim|adığım|ediğim)/i,
  /araç\s+çağrısı\s+yapama/i,
  /(?:dosya|araç)\s+erişimi\s+(?:olan|etkin)(?:[^.\n]|\.(?=\S)){0,40}oturum/i,
  /kopyala(?:[^.\n]|\.(?=\S)){0,80}kayde[dt]/i,
];

const REMOTE_ARTIFACT_COMPLETION_PATTERNS: RegExp[] = [
  /sandbox:\/mnt\/data\/[^\s)\]]+/i,
  /https?:\/\/[^\s)\]]*asyncgw\.teams\.microsoft\.com\/[^\s)\]]+\.(?:patch|diff)(?:[?#][^\s)\]]*)?/i,
  /\b(?:updated|modified|replaced|rewrote|saved|applied|prepared|created)\b[\s\S]{0,600}https?:\/\/[^\s)\]]*asyncgw\.teams\.microsoft\.com\/[^\s)\]]*\/views\/original\//i,
  /https?:\/\/[^\s)\]]*asyncgw\.teams\.microsoft\.com\/[^\s)\]]*\/views\/original\/[\s\S]{0,600}\b(?:updated|modified|replaced|rewrote|saved|applied|prepared|created)\b/i,
  /\b(?:updated|modified|replaced|rewrote|saved|applied)\b[\s\S]{0,180}\uE200cite\uE202turn\d+file\d+\uE201/i,
  /\uE200cite\uE202turn\d+file\d+\uE201[\s\S]{0,180}\b(?:updated|modified|replaced|rewrote|saved|applied)\b/i,
];

const HALLUCINATED_COMPLETION_PATTERNS: RegExp[] = [
  /\b(?:requested\s+)?(?:local\s+)?(?:edit|update|change|modification)\s+(?:is|was)\s+(?:now\s+)?complete\b/i,
  /\bI(?:'ve|\s+have|\s+just|\s+now)?\s+(?:created|wrote|written|replaced|updated|saved|applied|added|overwrote|modified|generated|implemented|rewrote)\b/i,
  /\b(?:the\s+)?(?:file|readme|script|config|change|version|content)\s+(?:has|have|is|was|were)\s+(?:been\s+)?(?:created|replaced|updated|saved|written|applied|added|modified|overwritten)\b/i,
  /\bhere'?s\s+(?:the\s+)?(?:updated|new|simplified|replaced|final)\s+(?:file|readme|version|content)\b/i,
  /\b(?:created|wrote|written|generated|saved|added|produced|implemented|overwrote)\b[^.\n]{0,60}\b[\w-]{2,}\.[a-z]{1,4}\b/i,
  /\b(?:executed|ran|invoked|launched|compiled)\b[^.\n]{0,40}\b(?:it|them|this|the\s+(?:script|program|file|code|command|tests?)|python3?|node|\S{2,}\.[a-z]{1,4})\b/i,
];

export function looksLikeHallucinatedCompletion(text: string | null): boolean {
  if (!text) return false;
  const t = text.trim();
  if (t.length < 8) return false;
  return HALLUCINATED_COMPLETION_PATTERNS.some((re) => re.test(t));
}

export function looksLikeRemoteArtifactCompletion(text: string | null): boolean {
  if (!text) return false;
  const t = text.trim();
  if (t.length < 12) return false;
  return REMOTE_ARTIFACT_COMPLETION_PATTERNS.some((re) => re.test(t));
}

export function looksLikeConfabulation(text: string | null): boolean {
  if (!text) return false;
  const t = text.trim();
  if (t.length < 12) return false;
  return CONFABULATION_PATTERNS.some((re) => re.test(t));
}

const SELF_WRITTEN_RESULT = /<tool_(?:response|result)\b/i;
const ANY_FENCE = /```[A-Za-z0-9_.-]+[ \t]*\r?\n[\s\S]*?```/;

export function truncateAtFabricatedToolResponse(text: string, tools?: ToolDef[]): string {
  const i = fabricatedToolResponseAt(text, tools);
  return i < 0 ? text : text.slice(0, i).trimEnd();
}

export function fabricatedToolResponseAt(text: string, tools?: ToolDef[]): number {
  const i = text.search(SELF_WRITTEN_RESULT);
  if (i < 0) return -1;
  const head = text.slice(0, i);
  const acted = tools && tools.length > 0 ? parseToolCalls(head, tools).hasToolCalls : ANY_FENCE.test(head);
  return acted ? i : -1;
}

export function isProseDocument(parsed: ParseResult, text?: string, tools?: ToolDef[]): boolean {
  if (!parsed.hasToolCalls || parsed.toolCalls.length < 2) return false;
  if (text !== undefined && tools && tools.length > 0) {
    const first = findFirstToolFence(text, buildSpecMap(tools));
    if (first && !looksLikeDocumentPreamble(text.slice(0, first.start))) return false;
  }
  const prose = parsed.textContent ? parsed.textContent.trim() : "";
  const hasMarkdownHeaders = /^#{1,6}\s/m.test(prose);
  return parsed.toolCalls.length >= 4 || hasMarkdownHeaders || prose.length >= 300;
}

function looksLikeDocumentPreamble(preamble: string): boolean {
  return /^#{1,6}\s/m.test(preamble) || preamble.includes("```") || preamble.trim().length >= 200;
}

export function textAfterFirstToolCall(text: string, tools?: ToolDef[]): string {
  if (!tools || tools.length === 0) return "";
  const first = findFirstToolFence(text, buildSpecMap(tools));
  return first ? text.slice(first.end).trim() : "";
}

export function parseToolCalls(text: string, tools?: ToolDef[]): ParseResult {
  const specMap = tools && tools.length > 0 ? buildSpecMap(tools) : null;
  if (specMap) {
    const { calls, leftover } = parseFencedToolCalls(text, specMap);
    if (calls.length > 0) {
      return { hasToolCalls: true, toolCalls: calls, textContent: cleanLooseText(leftover) };
    }
  }

  const resolveName = (raw: unknown): string | undefined =>
    typeof raw === "string" ? (specMap?.get(raw)?.name ?? raw) : undefined;

  const toolCalls: ParsedToolCall[] = [];

  const jsonRegex = new RegExp(TOOL_CALL_REGEX.source, "g");
  let match: RegExpExecArray | null;

  while ((match = jsonRegex.exec(text)) !== null) {
    try {
      const parsed = JSON.parse(match[0]);
      const name = resolveName(parsed.tool);
      if (name) {
        toolCalls.push({
          id: `call_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`,
          type: "function",
          function: {
            name,
            arguments: typeof parsed.arguments === "string"
              ? parsed.arguments
              : JSON.stringify(parsed.arguments ?? {}),
          },
        });
      }
    } catch {
      log.error("Failed to parse tool call JSON:", match[0]);
    }
  }

  if (toolCalls.length === 0) {
    const fencedRegex = new RegExp(FENCED_TOOL_CALL_REGEX.source, "g");
    while ((match = fencedRegex.exec(text)) !== null) {
      try {
        const parsed = JSON.parse(match[1]);
        const name = resolveName(parsed.tool || parsed.name);
        if (name) {
          toolCalls.push({
            id: `call_${crypto.randomUUID().replace(/-/g, "").slice(0, 24)}`,
            type: "function",
            function: {
              name,
              arguments: typeof parsed.arguments === "string"
                ? parsed.arguments
                : JSON.stringify(parsed.arguments ?? {}),
            },
          });
        }
      } catch {
        log.error("Failed to parse fenced tool call JSON:", match[1]);
      }
    }
  }

  if (toolCalls.length === 0) {
    return { hasToolCalls: false, toolCalls: [], textContent: cleanLooseText(text) };
  }

  let remaining = text
    .replace(jsonRegex, "")
    .replace(new RegExp(FENCED_TOOL_CALL_REGEX.source, "g"), "")
    .replace(CONFIDENCE_REGEX, "")
    .replace(FINAL_OBJECT_REGEX, "")
    .replace(/```(?:json|tool_call)?\s*```/g, "")
    .replace(/```(?:json|tool_call)?/g, "")
    .trim();

  return {
    hasToolCalls: true,
    toolCalls,
    textContent: remaining || null,
  };
}
