import {
  ModelSession,
  type ModelSessionOptions,
  createLogger,
  trunc,
  getToneForModel,
  requestUsesAgent,
  isPremiumOnlyModel,
  noteAgentRouteDead,
  noteAgentRouteAlive,
  isAgentRouteAlive,
  PREMIUM_ONLY_AGENT_TONES,
  defaultFramingForModel,
  currentFramingVariant,
  transcriptStyleForVariant,
  priorityAccessExhaustionOf,
  notePriorityAccessExhausted,
  activePriorityAccessExhaustion,
  meteredAllowance,
  METERED_BUDGETS,
  type MeteredBudget,
  meteredBudgetOf,
  priorityAccessFallbackModel,
  PRIORITY_ACCESS_FALLBACK_ENV,
  couldBePriorityAccessPrefix,
  secondsUntilReset,
  type PriorityAccessExhaustion,
  formatMessages,
  imageUrls,
  parseToolCalls,
  looksLikeConfabulation,
  looksLikeHallucinatedCompletion,
  looksLikeRemoteArtifactCompletion,
  truncateAtFabricatedToolResponse,
  fabricatedToolResponseAt,
  textAfterFirstToolCall,
  longHeredocAsWrite,
  formatToolResponse,
  isProseDocument,
  getMessageContent,
  noteRequestOutcome,
  awaitDegradationBackoff,
  awaitTurnBudget,
  TurnBudgetHold,
  isPersonalMode,
  getImageArtifactToken,
  fetchImageBytes,
  type CapturedImage,
  type ParseResult,
} from "@m365-copilot/core";
import { ChatCompletionRequest } from "./schemas.js";
import type { z } from "zod/v4";

const log = createLogger("handler");

async function renderImagesMarkdown(images: CapturedImage[]): Promise<string> {
  if (images.length === 0) return "";
  let artifactToken: string | null = null;
  try { artifactToken = await getImageArtifactToken(); } catch (e: any) { log.info(`image token failed: ${e.message}`); }
  const parts: string[] = [];
  for (const img of images) {
    const url = img.referenceUrls[0];
    if (!url) continue;
    if (artifactToken) {
      try {
        const { data, contentType } = await fetchImageBytes(url, artifactToken);
        parts.push(`![generated image](data:${contentType};base64,${data.toString("base64")})`);
        continue;
      } catch (e: any) { log.info(`image fetch failed: ${e.message}`); }
    }
    parts.push(`![generated image](${url})`);
  }
  return parts.join("\n\n");
}

const CONFAB_FORCE_PROMPT =
  "The working directory and the files named in the task ARE present on a real filesystem right now. Do NOT ask me to paste anything, and do NOT say commands return no output — you have not run any command yet. Emit ONE ```bash block this turn: run `ls -la` and `cat` the relevant files. Output only the ```bash block, nothing else.";

const MIDTASK_FORCE_PROMPT =
  "Your tools are working: your previous command really ran on this machine, and its real output is in the last <tool_response>. Continue the task from there with ONE tool block for the next step. A missing program or library is not a missing tool — install it with the package manager (pip, npm, …) or do the job with what is already installed. Output only the tool block, nothing else.";

const HALLUCINATION_FORCE_PROMPT =
  "You have NOT actually done that — no tool ran this turn, so nothing changed on disk. Do not claim a file was created, replaced, or updated until a <tool_response> confirms it. Emit ONE ```bash block now that performs the change for real (write the file with a `cat > path <<'EOF' … EOF` heredoc), and nothing else.";

const REMOTE_ARTIFACT_FORCE_PROMPT =
  "The patch or download link you produced exists only in M365's remote environment and is NOT a file in the caller's working directory. Do NOT create, download, or apply a patch, and do NOT use a Teams artifact link. Use the provided local edit/write tool directly; if needed, emit ONE ```bash block that modifies the named local file in place. Output only that single local tool call, nothing else.";

const OUTPUT_CHAR_CEILING = process.env.M365_OUTPUT_CHAR_CEILING !== undefined
  ? Number(process.env.M365_OUTPUT_CHAR_CEILING)
  : 12_000;

function outputFinishReason(text: string): "stop" | "length" {
  if (OUTPUT_CHAR_CEILING > 0 && text.length >= OUTPUT_CHAR_CEILING) {
    log.info(`Output at ceiling (${text.length} ≥ ${OUTPUT_CHAR_CEILING} chars) — finish_reason=length (likely truncated; harness should continue)`);
    return "length";
  }
  return "stop";
}

const DEFAULT_EARLY_STOP_AFTER = 2_000;
function earlyStopAfter(): number {
  const v = Number(process.env.M365_EARLY_STOP_AFTER);
  return process.env.M365_EARLY_STOP_AFTER !== undefined && Number.isFinite(v) && v >= 0 ? v : DEFAULT_EARLY_STOP_AFTER;
}

const UNSTREAMED_BUDGET_WAIT_LIMIT_MS = 2 * 60_000;

type ChatBody = z.infer<typeof ChatCompletionRequest>;
type ParsedMessage = ChatBody["messages"][number];

interface ConversationState {
  session: ModelSession;
  sentMessageCount: number;
  lastAccessedAt: number;
  pendingNote: string | null;
  lastSentIsReset: boolean;
  servedModel?: string;
  accountId: string | null;
}

const TRAILING_NOTE_CHARS = 120;
export function executedOnlyFirstNote(fabricated: boolean, droppedCalls: number, trailingChars = 0): string | null {
  if (fabricated) {
    return "(Note: only the first tool call in your previous reply was actually run. Everything you wrote after it, including the <tool_response> you wrote yourself, did not happen. Here is the real output of that first call:)";
  }
  const speculated = trailingChars >= TRAILING_NOTE_CHARS;
  if (droppedCalls > 0) {
    const rest = droppedCalls === 1 ? "the other one was not" : `the other ${droppedCalls} were not`;
    const tail = speculated ? " Anything you wrote after it was written before its result existed." : "";
    return `(Note: only the first of the ${droppedCalls + 1} tool calls in your previous reply was run; ${rest}.${tail} Here is the real output of the first one:)`;
  }
  if (speculated) {
    return "(Note: only the tool call in your previous reply was acted on; anything you wrote after it was written before its result existed. Here is the real output of that call:)";
  }
  return null;
}

const DEFAULT_SESSION_IDLE_MS = 60 * 60 * 1000;

export class SessionPool {
  static idleMs(): number {
    const env = Number(process.env.M365_SESSION_IDLE_MS);
    return Number.isFinite(env) && env > 0 ? env : DEFAULT_SESSION_IDLE_MS;
  }

  private conversations = new Map<string, ConversationState>();
  private sessionOptions: ModelSessionOptions;

  constructor(sessionOptions: ModelSessionOptions = {}) {
    this.sessionOptions = sessionOptions;
  }

  resolve(messages: ParsedMessage[], tools?: ChatBody["tools"], accountId?: string | null): ConversationState {
    this.evictStale();

    const account = accountId ?? null;
    const fingerprint = this.fingerprint(messages, tools, account);
    const existing = this.conversations.get(fingerprint);

    if (existing) {
      const before = existing.sentMessageCount;
      if (messages.length <= before && !existing.lastSentIsReset) {
        log.info(`Conversation ${fingerprint}: not a continuation (${messages.length} <= ${before} sent), starting a fresh M365 conversation`);
        existing.session.newConversation();
        existing.sentMessageCount = 0;
        existing.lastSentIsReset = true;
        existing.pendingNote = null;
      } else {
        existing.lastAccessedAt = Date.now();
      }
      return existing;
    }

    log.info(`New conversation ${fingerprint}${account ? ` account=${account}` : ""}, ${this.conversations.size} active`);
    const state: ConversationState = {
      session: new ModelSession({ ...this.sessionOptions, accountId: account }),
      sentMessageCount: 0,
      lastSentIsReset: false,
      pendingNote: null,
      lastAccessedAt: Date.now(),
      accountId: account,
    };
    this.conversations.set(fingerprint, state);
    return state;
  }

  private fingerprint(messages: ParsedMessage[], tools?: ChatBody["tools"], accountId?: string | null): string {
    const firstUser = messages.find(m => m.role === "user");
    const text = firstUser ? getMessageContent(firstUser) : "";
    const toolNames = [...new Set((tools ?? []).map(t => t.function.name))].sort().join(",");
    return simpleHash(`${accountId ?? ""}\n${toolNames}\n${text}`);
  }

  liveAccountIds(): string[] {
    const ids = new Set<string>();
    for (const state of this.conversations.values()) {
      if (state.accountId !== null) ids.add(state.accountId);
    }
    return [...ids];
  }

  forgetAccount(accountId: string): number {
    let dropped = 0;
    for (const [key, state] of this.conversations) {
      if (state.accountId === accountId) {
        this.conversations.delete(key);
        dropped++;
      }
    }
    if (dropped) log.info(`Dropped ${dropped} conversation(s) of account=${accountId}`);
    return dropped;
  }

  private evictStale() {
    const now = Date.now();
    const idleMs = SessionPool.idleMs();
    for (const [key, state] of this.conversations) {
      if (now - state.lastAccessedAt > idleMs) {
        log.info(`Evicting idle conversation ${key} (idle > ${Math.round(idleMs / 1000)} s)`);
        this.conversations.delete(key);
      }
    }
  }

  get size(): number {
    return this.conversations.size;
  }
}

function simpleHash(str: string): string {
  let hash = 0;
  for (let i = 0; i < str.length; i++) {
    hash = ((hash << 5) - hash + str.charCodeAt(i)) | 0;
  }
  return String(hash);
}

function parseReply(text: string, tools: Parameters<typeof parseToolCalls>[1]): { parsed: ParseResult; judged: string | null } {
  const parsed = parseToolCalls(text, tools);
  if (isProseDocument(parsed, text, tools)) {
    log.info(`Response is a prose document (${parsed.toolCalls.length} embedded fences), returning as text instead of executing`);
    const opening = text.split("```")[0].trim().split(/\n\s*\n/)[0];
    return { parsed: { hasToolCalls: false, toolCalls: [], textContent: text }, judged: opening };
  }
  return { parsed, judged: parsed.textContent };
}

export function formatDeltaMessages(messages: ParsedMessage[], history: ParsedMessage[] = messages): string {
  const parts: string[] = [];
  for (const m of messages) {
    if (m.role === "assistant") {
      continue;
    } else if (m.role === "tool") {
      parts.push(formatToolResponse(m, history));
    } else if (m.role === "system") {
    } else {
      parts.push(`<${m.role}>\n${getMessageContent(m)}\n</${m.role}>`);
    }
  }
  return parts.join("\n\n");
}

export async function handleChatCompletion(
  body: ChatBody,
  pool: SessionPool,
  opts: {
    signal?: AbortSignal;
    ensureCredentials?: (accountId: string | null) => Promise<void>;
    accountId?: string | null;
  } = {},
): Promise<Response> {
  const account = opts.accountId ?? null;
  const conv = pool.resolve(body.messages, body.tools, account);
  const { session } = conv;
  const hasTools = body.tools && body.tools.length > 0 && body.tool_choice !== "none";
  const requestedModel = body.model;

  // Metering is a property of the paid scenarios; a rule of the upstream API.
  // Consumer sessions enforce their own CostQuota upstream.
  const metered = !isPersonalMode();
  const requestedBudget = metered ? meteredBudgetOf(requestedModel) : null;
  const knownExhaustion = metered ? activePriorityAccessExhaustion(requestedBudget ?? requestedModel) : null;
  const fallbackModel = knownExhaustion && requestedBudget ? priorityAccessFallbackModel(requestedBudget) : null;
  if (knownExhaustion && !fallbackModel) {
    log.info(`Priority access exhausted (${knownExhaustion.window}) until ${knownExhaustion.resetsAt.toISOString()} — 429 without contacting M365`);
    return priorityAccessResponse(knownExhaustion, requestedModel);
  }
  let model = fallbackModel ?? requestedModel;
  if (fallbackModel) log.info(`Priority access exhausted until ${knownExhaustion!.resetsAt.toISOString()} — serving ${requestedModel} with ${fallbackModel} (${PRIORITY_ACCESS_FALLBACK_ENV[requestedBudget!]})`);
  if (conv.servedModel && conv.servedModel !== model) {
    log.info(`Conversation switches model ${conv.servedModel} → ${model}: fresh conversation, full history`);
    session.newConversation();
    conv.sentMessageCount = 0;
    conv.lastSentIsReset = true;
  }

  let tone = getToneForModel(model);
  let framingVariant = currentFramingVariant(defaultFramingForModel(model));
  let useToolAgent = requestUsesAgent(model, !!hasTools);

  const isFirstTurn = session.turnCount === 0;
  const sendFullHistory = isFirstTurn || conv.sentMessageCount === 0;
  const convId = session.conversationId;
  let text: string;
  let inputImages: string[];
  if (sendFullHistory) {
    text = formatMessages(body.messages, body.tools, body.tool_choice, convId, framingVariant);
    inputImages = imageUrls(body.messages);
    log.info(`Chat completion: model=${model}, stream=${body.stream}, messages=${body.messages.length}, turn=${session.turnCount}, mode=full, cid=${convId}`);
  } else {
    const newMessages = body.messages.slice(conv.sentMessageCount);
    inputImages = imageUrls(newMessages);
    const delta = newMessages.length > 0 ? formatDeltaMessages(newMessages, body.messages) : "";
    if (delta.length > 0) {
      text = delta;
      if (conv.pendingNote && newMessages.some((m) => m.role === "tool")) {
        text = `${conv.pendingNote}\n\n${delta}`;
        log.info("Prepended the only-the-first-call-ran note (last reply was partly executed)");
      }
      log.info(`Chat completion: model=${model}, stream=${body.stream}, messages=${body.messages.length}, new=${newMessages.length}, turn=${session.turnCount}, mode=delta, cid=${convId}`);
    } else {
      text = "Please continue.";
      log.info(`Chat completion: model=${model}, stream=${body.stream}, messages=${body.messages.length}, turn=${session.turnCount}, mode=retry, cid=${convId}`);
    }
  }
  conv.pendingNote = null;

  log.debug("Formatted prompt:", trunc(text, 1000));

  const completionId = `chatcmpl-${crypto.randomUUID()}`;
  const created = Math.floor(Date.now() / 1000);

  if (opts.ensureCredentials) {
    try {
      await opts.ensureCredentials(conv.accountId);
    } catch (err: any) {
      log.info(`Pre-flight failed: ${err?.message}`);
      return jsonResponse(401, {
        error: {
          message: err?.message ?? "authentication failed",
          type: "authentication_error",
          code: "m365_personal_session_dead",
        },
      });
    }
  }

  const MAX_RETRIES = 2;
  const SHORT_RETRY_DELAY_MS = 2_000;

  let lastThrottle: { current: number; max: number } | null = null;
  let lastContentOrigin: string | null | undefined;
  let lastMessageType: string | null | undefined;
  let lastScores: Record<string, number> | null | undefined;
  let lastTurnCount: number | null | undefined;
  let lastMetering: Record<string, number> | null | undefined;

  async function runBuffered(
    onDelta?: (delta: string) => void,
  ): Promise<{ fullText: string } | { error: Response }> {
    let agentRefreshed = false;
    let disengageRetried = false;
    let agentFallbackDone = false;
    let priorityFallbackDone = false;
    let internalErrorFreshRetried = false;
    let originalText = text;
    await awaitDegradationBackoff(conv.accountId);
    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      const earlyStop = hasTools && !process.env.M365_NO_EARLY_STOP ? new AbortController() : null;
      const stopAfter = earlyStopAfter();
      const signal = !earlyStop ? opts.signal
        : opts.signal ? AbortSignal.any([opts.signal, earlyStop.signal]) : earlyStop.signal;
      let copilotStream;
      try {
        const waitForBudget = () => awaitTurnBudget({ signal: opts.signal, maxWaitMs: body.stream ? undefined : UNSTREAMED_BUDGET_WAIT_LIMIT_MS });
        await waitForBudget();
        copilotStream = await session.run(text, model, signal, useToolAgent, inputImages, waitForBudget);
      } catch (err: any) {
        if (err instanceof TurnBudgetHold) return { error: turnBudgetHoldResponse(err) };
        return { error: jsonResponse(502, { error: { message: err.message, type: "upstream_error" } }) };
      }

      let fullText = "";
      let tagAt = -1;
      try {
        for await (const delta of copilotStream) {
          fullText += delta;
          onDelta?.(delta);
          if (!earlyStop || earlyStop.signal.aborted) continue;
          const sofar = (copilotStream.fullText?.length ?? 0) > fullText.length ? copilotStream.fullText! : fullText;
          if (tagAt < 0 && /<tool_/i.test(fullText.slice(Math.max(0, fullText.length - delta.length - 16)))) {
            tagAt = fabricatedToolResponseAt(sofar, body.tools);
          }
          if (tagAt >= 0 && sofar.length - tagAt >= stopAfter) {
            log.info(`Model is writing its own <tool_response> after its call — stopping the turn ${sofar.length} chars in (tag at ${tagAt})`);
            earlyStop.abort();
          }
        }
        if (copilotStream.fullText && copilotStream.fullText.length > fullText.length) {
          fullText = copilotStream.fullText;
        }
      } catch (err: any) {
        return { error: jsonResponse(502, { error: { message: err.message, type: "upstream_error" } }) };
      }

      const images = copilotStream.images ?? [];
      if (images.length > 0) {
        const imageMd = await renderImagesMarkdown(images);
        if (imageMd) {
          const addition = fullText.length > 0 ? `\n\n${imageMd}` : imageMd;
          fullText += addition;
          onDelta?.(addition);
        }
      }

      lastThrottle = copilotStream.throttle;
      lastContentOrigin = copilotStream.contentOrigin;
      lastMessageType = copilotStream.messageType;
      lastScores = copilotStream.scores;
      lastTurnCount = copilotStream.turnCount;
      lastMetering = copilotStream.metering;

      if (copilotStream.result?.value === "Throttled") {
        noteRequestOutcome(true, convId, conv.accountId);
        log.info(`Upstream Throttled (${copilotStream.result.errorCode ?? "no errorCode"}) — 429, no retry`);
        return { error: throttledResponse(copilotStream.result) };
      }

      if (
        useToolAgent && !agentFallbackDone &&
        PREMIUM_ONLY_AGENT_TONES.has(tone) &&
        !isAgentRouteAlive(tone) &&
        process.env.M365_FORCE_AGENT !== "1" &&
        copilotStream.result?.value === "InternalError" &&
        !copilotStream.hasContent && fullText.length === 0
      ) {
        agentFallbackDone = true;
        noteAgentRouteDead(tone);
        useToolAgent = false;
        session.newConversation();
        text = formatMessages(body.messages, body.tools, body.tool_choice, session.conversationId, framingVariant);
        inputImages = imageUrls(body.messages);
        originalText = text;
        log.info(`Agent route dead for ${tone} (InternalError) — this account isn't premium; re-sending agent-less with '${framingVariant}' framing in a fresh conversation`);
        attempt--;
        continue;
      }

      const exhausted = priorityAccessExhaustionOf(copilotStream.result, fullText);
      if (exhausted) {
        const budget = metered ? meteredBudgetOf(model) : null;
        if (metered && (budget || copilotStream.result?.value === "OutOfCredits")) notePriorityAccessExhausted(budget ?? model, exhausted);
        const fallback = budget && !priorityFallbackDone ? priorityAccessFallbackModel(budget) : null;
        if (fallback) {
          priorityFallbackDone = true;
          log.info(`Priority access exhausted (${exhausted.window}) — re-sending to ${fallback} (${PRIORITY_ACCESS_FALLBACK_ENV[budget!]}) in a fresh conversation`);
          model = fallback;
          tone = getToneForModel(model);
          framingVariant = currentFramingVariant(defaultFramingForModel(model));
          useToolAgent = requestUsesAgent(model, !!hasTools);
          session.newConversation();
          text = formatMessages(body.messages, body.tools, body.tool_choice, session.conversationId, framingVariant);
          inputImages = imageUrls(body.messages);
          originalText = text;
          attempt--;
          continue;
        }
        noteRequestOutcome(false, convId, conv.accountId);
        log.info(`Priority access exhausted (${exhausted.window}) — resets ${exhausted.resetsAt.toISOString()}`);
        return { error: priorityAccessResponse(exhausted, model) };
      }

      if (copilotStream.hasContent || fullText.length > 0) {
        noteRequestOutcome(false, convId, conv.accountId);
        if (useToolAgent) noteAgentRouteAlive(tone);
        conv.servedModel = model;
        inputImages = [];
        return { fullText };
      }

      // Disengagement is permanent; observed, undocumented by Microsoft.
      if (copilotStream.messageType === "Disengaged") {
        if (hasTools && !disengageRetried && !process.env.M365_NO_DISENGAGE_RETRY) {
          disengageRetried = true;
          session.newConversation();
          const retryVariant = transcriptStyleForVariant(framingVariant).framingTag === "system" ? "softened" : framingVariant;
          text = formatMessages(body.messages, body.tools, body.tool_choice, session.conversationId, retryVariant);
          inputImages = imageUrls(body.messages);
          log.info(`Upstream Disengaged — retrying once with '${retryVariant}' framing in a fresh conversation (F22)`);
          attempt--;
          continue;
        }
        log.info("Upstream Disengaged — failing fast (no retry) to preserve quota");
        return {
          error: jsonResponse(502, {
            error: {
              message: "M365 Copilot disengaged from this request (its safety filter declined to answer). Common causes: too many tools, jailbreak-shaped instructions, or pairing a non-default model with the tool agent. Reduce the toolset or use the default model.",
              type: "disengaged",
            },
          }),
        };
      }

      if (
        !internalErrorFreshRetried &&
        copilotStream.result?.value === "InternalError" &&
        !copilotStream.hasContent && fullText.length === 0 &&
        !isPremiumOnlyModel(model) &&
        attempt < MAX_RETRIES
      ) {
        internalErrorFreshRetried = true;
        session.newConversation();
        text = formatMessages(body.messages, body.tools, body.tool_choice, session.conversationId, framingVariant);
        inputImages = imageUrls(body.messages);
        originalText = text;
        log.info(`Upstream InternalError — re-sending the whole request in a fresh conversation (attempt ${attempt + 1}/${MAX_RETRIES})`);
        await new Promise(r => setTimeout(r, SHORT_RETRY_DELAY_MS));
        continue;
      }

      const t = copilotStream.throttle;
      if (t && t.current >= t.max) {
        return { error: rateLimitResponse(t) };
      }
      if (attempt < MAX_RETRIES) {
        if (!agentRefreshed) {
          agentRefreshed = true;
          const agentChanged = await session.refreshAgent();
          if (agentChanged) {
            log.info("Agent re-resolved after empty reply, resending original prompt");
            text = originalText;
            await new Promise(r => setTimeout(r, SHORT_RETRY_DELAY_MS));
            continue;
          }
        }
        log.info(`Empty upstream response, quick retry in ${SHORT_RETRY_DELAY_MS / 1000}s (attempt ${attempt + 1}/${MAX_RETRIES})`);
        await new Promise(r => setTimeout(r, SHORT_RETRY_DELAY_MS));
        text = "Please continue.";
      } else {
        if (isPremiumOnlyModel(model) && copilotStream.result?.value === "InternalError") {
          log.info(`${model}: InternalError on every attempt — premium-only model on a non-premium account?`);
          return { error: premiumOnlyModelResponse(model) };
        }
        noteRequestOutcome(true, convId, conv.accountId);
        return { error: emptyResponseResponse(t) };
      }
    }
    noteRequestOutcome(true, convId, conv.accountId);
    return { error: emptyResponseResponse(null) };
  }

  type Produced =
    | { kind: "error"; resp: Response }
    | { kind: "text"; text: string }
    | { kind: "tools"; toolCalls: ReturnType<typeof parseToolCalls>["toolCalls"] };

  async function produce(onDelta?: (delta: string) => void): Promise<Produced> {
  const markUsed = () => {
    conv.lastSentIsReset = false;
    conv.lastAccessedAt = Date.now();
  };
  if (hasTools) {
    let cutFabricated = false;
    let droppedCalls = 0;
    const stopAtFabricatedResult = (raw: string): string => {
      const cut = truncateAtFabricatedToolResponse(raw, body.tools);
      cutFabricated = cut !== raw;
      if (cutFabricated) log.info(`Model wrote its own <tool_response> — cut ${raw.length - cut.length} fabricated chars (stop sequence)`);
      return cut;
    };
    const result = await runBuffered();
    if ("error" in result) return { kind: "error", resp: result.error };
    conv.sentMessageCount = body.messages.length; markUsed();
    let fullText = stopAtFabricatedResult(result.fullText);

    log.debug("Raw response (tool mode):", trunc(fullText, 1000));
    let { parsed, judged } = parseReply(fullText, body.tools);
    log.info(`Parse result: hasToolCalls=${parsed.hasToolCalls}, count=${parsed.toolCalls.length}`);

    const maxConfabRetries = process.env.M365_NO_CONFAB_RETRY
      ? 0
      : Number(process.env.M365_CONFAB_RETRIES ?? 1);
    const everActed = (body.messages ?? []).some(
      (m) => m.role === "assistant" && Array.isArray(m.tool_calls) && m.tool_calls.length > 0,
    );
    for (let attempt = 0; attempt < maxConfabRetries && !parsed.hasToolCalls; attempt++) {
      const confab = looksLikeConfabulation(judged);
      const remoteArtifact = looksLikeRemoteArtifactCompletion(parsed.textContent);
      const halluc = !everActed && looksLikeHallucinatedCompletion(judged);
      if (!confab && !remoteArtifact && !halluc) break;
      const retryKind = remoteArtifact ? "Remote artifact completion" : confab ? "Confabulation" : "Hallucinated completion";
      log.info(`${retryKind} detected (no tool call) — forcing retry ${attempt + 1}/${maxConfabRetries}`);
      text = remoteArtifact ? REMOTE_ARTIFACT_FORCE_PROMPT
        : confab ? (everActed ? MIDTASK_FORCE_PROMPT : CONFAB_FORCE_PROMPT)
        : HALLUCINATION_FORCE_PROMPT;
      const retry = await runBuffered();
      if ("error" in retry) return { kind: "error", resp: retry.error };
      conv.sentMessageCount = body.messages.length; markUsed();
      fullText = stopAtFabricatedResult(retry.fullText);
      ({ parsed, judged } = parseReply(fullText, body.tools));
      log.info(`After forcing retry: hasToolCalls=${parsed.hasToolCalls}, count=${parsed.toolCalls.length}`);
    }

    if (!parsed.hasToolCalls && looksLikeRemoteArtifactCompletion(parsed.textContent)) {
      log.info("Remote-artifact mutation claim persisted without a local tool call after forcing retries — failing closed");
      return {
        kind: "error",
        resp: jsonResponse(502, {
          error: {
            message: "M365 returned a remote Teams or /mnt/data artifact instead of calling the local editing tools. No local file was changed. Retry with gpt-5.5-think-deeper, the recommended model for tool calling.",
            type: "file_mutation_without_local_tool",
          },
        }),
      };
    }

    if (parsed.hasToolCalls && parsed.textContent) {
      const extraText = parsed.textContent.trim();
      if (extraText.length > 0) {
        log.info(`Mixed output detected (${extraText.length} chars of text alongside ${parsed.toolCalls.length} tool calls), stripping text`);
        log.debug("Stripped text:", trunc(extraText, 500));
        parsed = { ...parsed, textContent: null };
      }
    }

    if (parsed.hasToolCalls) {
      const replyCall = parsed.toolCalls.find(tc => tc.function.name === "reply");
      const realToolCalls = parsed.toolCalls.filter(tc => tc.function.name !== "reply");

      if (replyCall && realToolCalls.length === 0) {
        let replyText: string;
        try {
          const args = JSON.parse(replyCall.function.arguments);
          replyText = args.text || args.message || args.content || fullText;
        } catch {
          replyText = fullText;
        }
        log.info("Reply tool detected, converting to text response");
        return { kind: "text", text: replyText };
      }

      if (realToolCalls.length > 0) {
        parsed.toolCalls = realToolCalls;
      }

      if (!process.env.M365_ALLOW_MULTI_TOOL && parsed.toolCalls.length > 1) {
        log.info(`One-call-per-turn: keeping ${parsed.toolCalls[0].function.name}, dropping ${parsed.toolCalls.length - 1} batched call(s)`);
        droppedCalls = parsed.toolCalls.length - 1;
        parsed.toolCalls = [parsed.toolCalls[0]];
      }
    }

    if (parsed.hasToolCalls && parsed.toolCalls.length > 0) {
      const asWrite = longHeredocAsWrite(parsed.toolCalls[0], body.tools);
      let rewriteNote: string | null = null;
      if (asWrite) {
        log.info(`Heredoc write to ${asWrite.path} is over the Windows command-line cap — sent as ${asWrite.call.function.name} instead`);
        parsed.toolCalls = [asWrite.call];
        rewriteNote = `(Note: your heredoc to \`${asWrite.path}\` was longer than the ~8,000-character Windows command-line limit and would have been cut off, so it was written with the \`${asWrite.call.function.name}\` tool instead.${asWrite.skipped ? ` The commands after the heredoc did NOT run: \`${asWrite.skipped.slice(0, 200)}\`.` : ""})`;
      }
      const notes = [rewriteNote, executedOnlyFirstNote(cutFabricated, droppedCalls, textAfterFirstToolCall(fullText, body.tools).length)].filter(Boolean);
      conv.pendingNote = notes.length ? notes.join("\n") : null;
      markUsed();
      return { kind: "tools", toolCalls: parsed.toolCalls };
    }
    markUsed();
    return { kind: "text", text: fullText };
  } else {
    const result = await runBuffered(onDelta);
    if ("error" in result) return { kind: "error", resp: result.error };
    conv.sentMessageCount = body.messages.length; markUsed();
    return { kind: "text", text: result.fullText };
  }
  }

  const includeUsage = !!body.stream_options?.include_usage;
  const usage = () => buildUsage(lastThrottle, lastContentOrigin, lastMessageType, lastScores, lastTurnCount, lastMetering);

  if (!body.stream) {
    const p = await produce();
    if (p.kind === "error") return p.resp;
    if (p.kind === "tools") {
      return jsonResponse(200, {
        id: completionId, object: "chat.completion", created, model,
        choices: [{ index: 0, message: { role: "assistant", content: null, tool_calls: p.toolCalls }, finish_reason: "tool_calls" }],
        usage: usage(),
      });
    }
    return jsonResponse(200, {
      id: completionId, object: "chat.completion", created, model,
      choices: [{ index: 0, message: { role: "assistant", content: p.text }, finish_reason: outputFinishReason(p.text) }],
      usage: usage(),
    });
  }

  return sseResponse(new ReadableStream({
    async start(controller) {
      const enc = new TextEncoder();
      const send = (obj: unknown) => controller.enqueue(enc.encode(`data: ${JSON.stringify(obj)}\n\n`));
      const base = { id: completionId, object: "chat.completion.chunk", created, model };
      send({ ...base, choices: [{ index: 0, delta: { role: "assistant" }, finish_reason: null }] });
      const hb = setInterval(() => {
        try { send({ ...base, choices: [{ index: 0, delta: {}, finish_reason: null }] }); } catch {}
      }, 15000);

      let sent = "";
      let head = "";
      let gated = true;
      const liveDelta = hasTools ? undefined : (delta: string) => {
        if (!delta) return;
        if (gated) {
          head += delta;
          if (couldBePriorityAccessPrefix(head)) return;
          gated = false;
          delta = head;
        }
        sent += delta;
        try { send({ ...base, choices: [{ index: 0, delta: { content: delta }, finish_reason: null }] }); } catch {}
      };

      let p: Produced;
      try { p = await produce(liveDelta); }
      catch (err: any) { p = { kind: "error", resp: jsonResponse(502, { error: { message: err?.message ?? "stream error", type: "upstream_error" } }) }; }
      clearInterval(hb);
      try {
        if (p.kind === "error") {
          let message = "upstream error";
          let type = "upstream_error";
          let code: string | undefined;
          try {
            const parsed = JSON.parse(await p.resp.text())?.error;
            if (parsed?.message) message = parsed.message;
            if (parsed?.type) type = parsed.type;
            if (parsed?.code) code = parsed.code;
          } catch {}
          const retryAfter = p.resp.headers.get("Retry-After");
          send({ ...base, error: { message, type, ...(code ? { code } : {}), ...(retryAfter ? { retry_after: Number(retryAfter) } : {}) } });
        } else if (p.kind === "tools") {
          p.toolCalls.forEach((tc, i) => {
            log.debug(`Tool call: ${tc.function.name} ${trunc(tc.function.arguments, 200)}`);
            send({ ...base, choices: [{ index: 0, delta: { tool_calls: [{ index: i, id: tc.id, type: "function", function: { name: tc.function.name, arguments: tc.function.arguments } }] }, finish_reason: null }] });
          });
          send({ ...base, choices: [{ index: 0, delta: {}, finish_reason: "tool_calls" }], ...(includeUsage ? { usage: usage() } : {}) });
        } else {
          const remainder = p.text.startsWith(sent) ? p.text.slice(sent.length) : "";
          if (!p.text.startsWith(sent)) log.info(`Streamed prefix diverged from final text (sent ${sent.length}, final ${p.text.length} chars) — not re-sending to avoid duplication`);
          if (remainder) send({ ...base, choices: [{ index: 0, delta: { content: remainder }, finish_reason: null }] });
          send({ ...base, choices: [{ index: 0, delta: {}, finish_reason: outputFinishReason(p.text) }], ...(includeUsage ? { usage: usage() } : {}) });
        }
      } catch {
      } finally {
        try { controller.enqueue(enc.encode("data: [DONE]\n\n")); controller.close(); } catch {}
      }
    },
  }));
}

function buildUsage(
  throttle: { current: number; max: number } | null,
  contentOrigin?: string | null,
  messageType?: string | null,
  scores?: Record<string, number> | null,
  turnCount?: number | null,
  metering?: Record<string, number> | null,
): Record<string, unknown> {
  const base: Record<string, unknown> = {
    prompt_tokens: 0,
    completion_tokens: 0,
    total_tokens: 0,
  };
  if (throttle) {
    base.x_m365_conversation_messages = throttle.current;
    base.x_m365_conversation_max = throttle.max;
    base.x_m365_conversation_pct = Math.min(100, Math.round((throttle.current / throttle.max) * 100));
    base.x_m365_conversation_remaining = Math.max(0, throttle.max - throttle.current);
  }
  if (contentOrigin) base.x_m365_content_origin = contentOrigin;
  if (messageType) base.x_m365_message_type = messageType;
  if (typeof turnCount === "number") base.x_m365_turn_count = turnCount;
  if (scores) {
    base.x_m365_classifier_scores = scores;
    if (typeof scores.dea_violation === "number") base.x_m365_dea_score = scores.dea_violation;
    if (typeof scores.BotOffense === "number") base.x_m365_offense_score = scores.BotOffense;
  }
  for (const [budget, { usageKey }] of Object.entries(METERED_BUDGETS)) {
    const left = meteredAllowance(metering, budget as MeteredBudget);
    if (left?.daily !== undefined) base[`x_m365_${usageKey}_daily_remaining`] = left.daily;
    if (left?.weekly !== undefined) base[`x_m365_${usageKey}_weekly_remaining`] = left.weekly;
  }
  return base;
}

function jsonResponse(status: number, body: unknown, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", ...headers },
  });
}

function priorityAccessResponse(exhaustion: PriorityAccessExhaustion, model: string): Response {
  const retryAfter = secondsUntilReset(exhaustion);
  const label = exhaustion.model ?? model;
  const when = exhaustion.window === "week" ? "this week" : "today";
  return jsonResponse(
    429,
    {
      error: {
        message:
          `M365 Copilot's priority access to ${label} is used up for ${when} ` +
          `(resets ${exhaustion.resetsAt.toISOString()}, i.e. midnight UTC` +
          `${exhaustion.window === "week" ? " on Monday" : ""}). ` +
          `Switch to another model (e.g. claude-sonnet or gpt-5.5-think-deeper) or wait. ` +
          `Upstream said: ${exhaustion.message}`,
        type: "rate_limit_error",
        code: "priority_access_exhausted",
        param: exhaustion.window,
      },
    },
    { "Retry-After": String(retryAfter) },
  );
}

function sseResponse(stream: ReadableStream): Response {
  return new Response(stream, {
    headers: { "Content-Type": "text/event-stream", "Cache-Control": "no-cache" },
  });
}

function rateLimitMessage(throttle: { current: number; max: number } | null): string {
  return throttle
    ? `M365 Copilot rate limited (${throttle.current}/${throttle.max} messages used). Please wait and try again.`
    : "M365 Copilot returned an empty response. You may be rate limited. Please wait and try again.";
}

function rateLimitResponse(throttle: { current: number; max: number } | null): Response {
  return jsonResponse(429, { error: { message: rateLimitMessage(throttle), type: "rate_limit_error" } });
}

function throttledResponse(result: { errorCode?: string; message?: string }): Response {
  const scope = result.errorCode ? ` (${result.errorCode})` : "";
  return jsonResponse(429, {
    error: {
      message: `M365 Copilot throttled this account${scope}: ${result.message ?? "too many requests"} ` +
        `Retrying immediately does not help; it is triggered by starting many new conversations in a short time.`,
      type: "rate_limit_error",
      code: "m365_throttled",
      ...(result.errorCode ? { param: result.errorCode } : {}),
    },
  });
}

function turnBudgetHoldResponse(hold: TurnBudgetHold): Response {
  const retryAfter = Math.max(1, Math.ceil(hold.waitMs / 1000));
  return jsonResponse(
    429,
    {
      error: {
        message: `Not sent to M365 Copilot (M365_AVOID_THROTTLING): ${hold.message}.`,
        type: "rate_limit_error",
        code: hold.reason === "throttle hold" ? "m365_throttle_hold" : "m365_turn_budget",
      },
    },
    { "Retry-After": String(retryAfter) },
  );
}

function emptyResponseResponse(throttle: { current: number; max: number } | null): Response {
  const detail = throttle ? ` (throttle ${throttle.current}/${throttle.max})` : "";
  return jsonResponse(502, {
    error: {
      message: `M365 Copilot returned an empty response${detail} — likely a content filter, an invalid agent/session, or a transient upstream error.`,
      type: "upstream_empty_response",
    },
  });
}

function premiumOnlyModelResponse(model: string): Response {
  return jsonResponse(502, {
    error: {
      message: `M365 Copilot didn't serve ${model}: it returned InternalError on every attempt. ${model} is served only on a premium (paid Microsoft 365 Copilot) account; on any other account use claude-sonnet.`,
      type: "upstream_error",
      code: "model_route_unavailable",
    },
  });
}
