import WebSocket from "ws";
import type { ImageAnnotation } from "../io/image-input.js";
import { appendFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import {
  SignalRHandshakeResponse,
  DeltaUpdate,
  MessageUpdate,
  ThrottlingUpdate,
  CompletionFrame,
  CloseFrame,
} from "./schemas.js";
import {
  getToneForModel,
  type CopilotStream,
  type CapturedImage,
} from "../routing/copilot.js";
import {
  parseActionConfirmation,
  buildResumeInvokeAction,
  shouldAutoConfirm,
  ACTION_ALLOWED_MESSAGE_TYPES,
} from "../tools/native-actions.js";
import { createLogger, trunc, FRAME_DIR } from "../io/log.js";
import { recordThrottle, recordTurn } from "../budget/turn-budget.js";
import { parseMetering, meteredAllowance, METERED_BUDGETS, type MeteredBudget } from "../budget/priority-access.js";
import type { TurnContext } from "../turn-context.js";

const RS = "\x1E";
const log = createLogger("session");

// The exact frame the real m365.cloud.microsoft "Stop generating" button sends
// (captured June 2026, scripts/cancel-frame-capture.mjs). It's a normal type:1
// invocation with target "stop" and invocationId "1" (distinct from chat's "0"),
// sent on the same socket; the server acks with a type:3 completion and discards
// the partial answer. See docs/m365-copilot-api.md §6 and hypotheses.md F11.
const STOP_FRAME = JSON.stringify({ arguments: [{}], invocationId: "1", target: "stop", type: 1 }) + RS;

// Enabling these optionsSets unlocks M365's real server-side Python sandbox:
// the model writes and EXECUTES Python and returns true results (verified with a
// SHA-256 oracle — docs/hypotheses.md §8.9 / scripts/code-interpreter-probe.mjs).
// Applied only on the agent-less (plain-chat) path: the tool-calling agent wants
// the model to emit tool JSON, not run its own Python, so we leave that path
const CODE_INTERPRETER_OPTIONS_SETS = [
  "cwc_code_interpreter",
  "cwc_code_interpreter_amsfix",
  "cwc_code_interpreter_citation_fix",
  "code_interpreter_interactive_charts",
  "code_interpreter_matplotlib_patching",
];

// The image-generation optionsSets, lifted verbatim from the official web client
// capture (§14). `flux_v3` is BizChat's orchestration codename, NOT the model —
// artifacts come back DallE-named. The GUI sends this WHOLE set on every turn and
// lets the model pick the type from the prompt (standard / icon / story / designer
// dimensions), which is why "draw me X" works with no UI toggle — so we match it:
// send everything the GUI enables. Excludes only the GPT-V/upload family, which is
// image *input* (a separate capability), not generation. `…non_watermarked_storage`
const IMAGE_GEN_OPTIONS_SETS = [
  "cwc_flux_image",
  "cwc_flux_v3",
  "enable_gg_gpt",
  "flux_v3_progress_messages",
  "flux_v3_image_gen_enable_dimensions",
  "flux_v3_image_gen_enable_icon_dimensions",
  "flux_v3_image_gen_enable_story",
  "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
  "flux_v3_image_gen_enable_system_text_with_params",
  "flux_v3_image_gen_enable_non_watermarked_storage",
];

const ENTERPRISE_MESSAGE_TYPES = [
  "Chat",
  "Suggestion",
  "InternalSearchQuery",
  "Disengaged",
  "InternalLoaderMessage",
  "Progress",
  "RenderCardRequest",
  "SemanticSerp",
  "GenerateContentQuery",
  "SearchQuery",
  "ConfirmationCard",
  "DeveloperLogs",
  "EndOfRequest",
  "ReferencesListComplete",
  "GeneratedCode",
];

// --- Optional per-request frame dumping for reverse engineering ---
// Enabled by M365_DUMP_FRAMES=1. Every SignalR frame received is appended to a
// per-request NDJSON file under FRAME_DIR (~/.config/opencode-m365/frames/, or
// M365_FRAME_DIR). Cheap to run in production; gives us forensic data when M365
// starts emitting new fields.
const DUMP_FRAMES = !!process.env.M365_DUMP_FRAMES;
function dumpFrame(requestId: string, parsed: unknown, direction: "recv" | "send") {
  if (!DUMP_FRAMES) return;
  try {
    mkdirSync(FRAME_DIR, { recursive: true });
    appendFileSync(
      join(FRAME_DIR, `${requestId}.ndjson`),
      JSON.stringify({ t: Date.now(), dir: direction, frame: parsed }) + "\n",
    );
  } catch {
    // best effort
  }
}

/**
 * Fold one incoming piece of streamed text — a token DeltaUpdate (passed as
 * `answer + writeAtCursor`) or a full-text MessageUpdate snapshot — into the
 * running answer, returning the new answer and the suffix to stream (if any).
 *
 * M365 mixes token deltas with full-text snapshots, and the first token often
 * arrives ONLY as a snapshot, so naive delta concatenation drops the head. The
 * rules:
 *  - `next` no longer than the current answer → no growth, emit nothing.
 *  - `next` extends the current answer → advance and emit the appended suffix.
 *  - `next` is longer but DIVERGES (doesn't start with the current answer, e.g. a
 *    snapshot carrying a head the deltas never sent) → adopt it as authoritative
 *    for the buffered result but emit nothing: already-streamed bytes can't be
 *    retracted, so we never stream a non-prefix (which would dup/corrupt).
 * This keeps every emitted suffix a true prefix of the final answer.
 */
export function foldStreamText(
  answer: string,
  next: string,
): { answer: string; emit: string | null } {
  if (next.length <= answer.length) return { answer, emit: null };
  if (next.startsWith(answer)) return { answer: next, emit: next.slice(answer.length) };
  return { answer: next, emit: null };
}

/** The messageId a streaming `cursor` points at: `$['<id>'].adaptiveCards[0]…`. */
export function cursorMessageId(j: string | undefined): string | null {
  const m = j?.match(/^\$\['([^']+)'\]/);
  return m ? m[1] : null;
}

/**
 * Reassemble one turn's text from M365's PER-MESSAGE stream.
 *
 * A turn can carry several bot messages (distinct messageIds): Claude narrates
 * in one and writes its tool fence in the next, and Sonnet 5 interleaves text
 * with its own tool calls. Each message streams the same way — its first token
 * usually arrives only as a snapshot that carries a `cursor` naming the message,
 * then `writeAtCursor` deltas that carry NO messageId and extend whatever the
 * cursor last named, then a final snapshot.
 *
 * Folding all of that into ONE string (what foldStreamText alone did) loses the
 * head of every message after the first: the new message's head snapshot is
 * shorter than the accumulated answer, so it is dropped, and its deltas get
 * glued onto the previous message — `…the SECRET_CODE.bash\ngrep…`, a fence
 * with no opening backticks, so the tool call never parses (#29). Every
 * multi-message turn in the Sep 28 dumps was corrupted this way.
 *
 * Messages are kept separately and composed in order of first appearance,
 * separated by a blank line (they are separate bubbles in the real client).
 */
export class TurnTextComposer {
  private texts = new Map<string, string>();
  private target: string | null = null;

  /** A cursor frame: subsequent deltas extend message `id`. */
  cursor(id: string): void {
    this.target = id;
  }

  /** A full-text snapshot of one message. Before any cursor has been seen, the
   *  snapshot's message also becomes the delta target (older frame shapes). */
  snapshot(id: string | undefined, text: string): void {
    const key = id ?? this.target ?? "";
    if (this.target === null) this.target = key;
    this.texts.set(key, foldStreamText(this.texts.get(key) ?? "", text).answer);
  }

  /** A token delta for the message the cursor last named. */
  delta(text: string): void {
    const key = this.target ?? "";
    this.texts.set(key, (this.texts.get(key) ?? "") + text);
  }

  get text(): string {
    return [...this.texts.values()].filter((t) => t.length > 0).join("\n\n");
  }
}

/**
 * Attach a native custom action/plugin to a conversation and drive the confirm→invoke
 * round-trip over the WS (docs/hypotheses.md §12.6, H-NATIVE-6/7). All fields are
 * best-effort passthroughs modelled on the decompiled client; the exact inline-def
 * schema still wants a live `SEND_MSG=1` capture, so run with M365_DUMP_FRAMES=1 to
 * see what the server actually does and refine.
 */
export interface NativeActionConfig {
  /** Inline agent definitions (`gptDefinitions[]`) — the no-sideload agent attach. */
  gptDefinitions?: unknown[];
  /** `clientOverrides.capabilities[]` entries (e.g. {name:"RegisteredPlugins", plugins:[…]}). */
  capabilities?: unknown[];
  /** `plugins[]` entries ({Id, Source, Data?}). Merged even when an agent is attached. */
  plugins?: unknown[];
  /** Approve consequential (state-mutating) actions too. Default: only read-only ones. */
  autoConfirmAll?: boolean;
}

/** Per-turn chat options. */
export interface ChatTurnOptions {
  /** Request server-side image generation for this turn (§14). Adds the flux
   *  optionsSets + the GenerateGraphicArt allowedMessageType; the generated
   *  images surface on `stream.images`. Agent-less only. */
  generateImages?: boolean;
  /** Uploaded images this turn's message points at (image-input.ts). */
  images?: ImageAnnotation[];
}

export interface CopilotSessionOptions {
  agentId?: string;
  /** Reuse an existing session ID across reconnections. */
  sessionId?: string;
  /** Reuse an existing conversation ID so M365 finds the same server-side conversation. */
  conversationId?: string;
  /**
   * Ask M365 for a temporary chat by adding `disableMemory=1` to the Chathub URL.
   * Temporary chats retain context for their live ConversationId but are not
   * added to the Copilot history sidebar. Default: true.
   */
  temporaryChat?: boolean;
  /** Enable native custom-action support (H-NATIVE-6/7). Off = unchanged behaviour. */
  nativeActions?: NativeActionConfig;
  /**
   * The credential this conversation runs on, decided once by the caller
   * (`selectCredential` in app/credential.ts). The session never reads an auth
   * mode and never resolves an account: it carries `accountId`, `personal` and
   * the personal `record` the caller chose.
   */
  turnContext: (plan: TurnPlan) => TurnContext;
  accountId?: string | null;
  model?: string;
}

export interface TurnLiveOptions {
  generateImages?: boolean;
  images?: ImageAnnotation[];
  nativeActions?: boolean;
}

export interface TurnPlan {
  requestId: string;
  sessionId: string;
  conversationId: string;
  model: string;
  temporaryChat: boolean;
  agentId: string | undefined;
  nativeActions: boolean;
  images: boolean;
}

/**
 * A persistent conversation session with M365 Copilot.
 * Reuses the same sessionId/conversationId across turns,
 * reconnecting the WebSocket for each message.
 */
export class CopilotSession {
  private sessionId: string;
  private conversationId: string;
  private _turnCount = 0;
  private agentId?: string;
  private nativeActions?: NativeActionConfig;
  private temporaryChat: boolean;
  private turnContext: (plan: TurnPlan) => TurnContext;
  private accountId: string | null;
  private model: string;

  constructor(options: CopilotSessionOptions) {
    this.sessionId = options.sessionId ?? crypto.randomUUID();
    this.conversationId = options.conversationId ?? crypto.randomUUID();
    this.agentId = options.agentId;
    this.nativeActions = options.nativeActions;
    this.temporaryChat = options.temporaryChat ?? true;
    this.turnContext = options.turnContext;
    this.accountId = options.accountId ?? null;
    this.model = options.model ?? "m365-copilot";
    log.info(`New session: sid=${this.sessionId}, cid=${this.conversationId}, agent=${this.agentId ?? "none"}, nativeActions=${!!this.nativeActions}, temporaryChat=${this.temporaryChat}, account=${this.accountId ?? "none"}`);
  }

  /** Number of turns completed in this session */
  get turnCount(): number {
    return this._turnCount;
  }

  get account(): string | null {
    return this.accountId;
  }

  /**
   * Send a message in this conversation and stream the response.
   * Each turn opens a fresh WebSocket with invocationId "0" (per SignalR protocol).
   * Session/conversation IDs are reused so M365 maintains server-side context.
   */
  chat(_token: string, text: string, model: string = this.model, signal?: AbortSignal, opts?: ChatTurnOptions): Promise<CopilotStream> {
    this._turnCount++;
    const isFirst = this._turnCount === 1;
    const accountId = this.accountId;
    const accountTag = accountId ? ` account=${accountId}` : "";

    log.info(`Chat turn ${this._turnCount - 1}: model=${model}, isFirst=${isFirst}, text=${JSON.stringify(trunc(text, 200))}${accountTag}`);
    recordTurn(undefined, accountId);

    const wantImages = opts?.generateImages ?? false;
    const images = opts?.images ?? [];
    const requestId = crypto.randomUUID();
    const tone = getToneForModel(model);
    const sessionId = this.sessionId;
    const conversationId = this.conversationId;
    const agentId = this.agentId;
    const nativeActions = this.nativeActions;
    const turnContext = this.turnContext({
      requestId,
      sessionId,
      conversationId,
      model,
      temporaryChat: this.temporaryChat,
      agentId,
      nativeActions: !!nativeActions,
      images: wantImages,
    });
    const supports = (capability: string) => !wantImages || turnContext.capabilities.includes(capability);

    return new Promise((resolve, reject) => {
      let answer = "";
      const composer = new TurnTextComposer();
      let receivedContent = false;
      let throttleInfo: { current: number; max: number } | null = null;
      let contentOrigin: string | null = null;
      let messageType: string | null = null;
      let messageId: string | null = null;
      // Highest per-component score across all messages this turn. The most
      // recent score isn't necessarily the most informative — the worst one is.
      const maxScores: Record<string, number> = {};
      let turnCountServer: number | null = null;
      let turnState: string | null = null;
      // The final type:2 item's `result` — `{value: "Success"}` normally, and
      // `{value: "Throttled", errorCode: "PerUserThrottled", message}` when the
      // account is rate-limited: the turn then carries NO content frames, only a
      // BotConnection apology inside that item (#35).
      let serverResult: { value: string; errorCode?: string; message?: string } | null = null;
      // The final item's `throttling.metering` (priority-access allowances, §24 F55).
      let metering: Record<string, number> | null = null;
      // Generated images captured this turn (§14). Keyed by fileToken so the
      // repeated progress snapshots for one image collapse to a single entry.
      const imagesByToken = new Map<string, CapturedImage>();
      // Native-action round-trip state (H-NATIVE-6). `baseArgs` is the sent chat
      // envelope's arguments[0], reused verbatim (minus `message`) to resume an
      // action. `sawAction` records that the model triggered a custom action this
      // turn; `actionResumed` guards against double-resuming the same trigger.
      let baseArgs: Record<string, unknown> | null = null;
      let sawAction = false;
      let actionResumed = false;

      // Delta pump. The queue/state live at Promise scope (NOT inside the
      // asyncIterator factory) so deltas that arrive between `resolve(stream)` and
      // the consumer starting to iterate are buffered instead of dropped — the
      // early-token loss that made true incremental streaming lossy. onDelta/onDone/
      // onError are wired once here, so they never depend on iteration having begun.
      const queue: string[] = [];
      let streamDone = false;
      let streamError: Error | null = null;
      let waiting: {
        resolve: (result: IteratorResult<string>) => void;
        reject: (err: Error) => void;
      } | null = null;

      const onDelta = (chunk: string) => {
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.resolve({ value: chunk, done: false });
        } else {
          queue.push(chunk);
        }
      };
      const onDone = () => {
        streamDone = true;
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.resolve({ value: undefined as any, done: true });
        }
      };
      const onError = (err: Error) => {
        streamError = err;
        streamDone = true;
        if (waiting) {
          const w = waiting;
          waiting = null;
          w.reject(err);
        }
      };

      // Pull any generated images off a bot message. The GraphicArt frame arrives
      // repeatedly as a Progress snapshot (status climbs to 2 = ready) and again in
      // the final type:2 item, so we upsert by fileToken and keep the readiest copy.
      const captureImages = (m: any) => {
        const list = m?.contentGenerationProgressList;
        if (!Array.isArray(list)) return;
        for (const entry of list) {
          const urls = entry?.ImageReferenceUrls;
          if (!Array.isArray(urls) || urls.length === 0) continue;
          const key = entry.fileToken ?? urls[0];
          const prev = imagesByToken.get(key);
          if (prev && (prev.status ?? 0) >= (entry.status ?? 0)) continue;
          imagesByToken.set(key, {
            referenceUrls: urls,
            fileToken: entry.fileToken,
            pollUrl: entry.pollUrl,
            size: entry.size,
            orientation: entry.orientation,
            status: entry.status,
          });
        }
      };

      // Fold a token delta OR a full-text snapshot into `answer` and stream the
      // newly-appended suffix (see foldStreamText for the prefix-safe rules).
      const advance = (next: string) => {
        const r = foldStreamText(answer, next);
        if (r.answer !== answer) {
          answer = r.answer;
          receivedContent = true;
        }
        if (r.emit) onDelta(r.emit);
      };

      const stream: CopilotStream = {
        get fullText() {
          return answer;
        },
        get hasContent() {
          return receivedContent;
        },
        get throttle() {
          return throttleInfo;
        },
        get contentOrigin() {
          return contentOrigin;
        },
        get messageType() {
          return messageType;
        },
        get messageId() {
          return messageId;
        },
        get scores() {
          return Object.keys(maxScores).length ? { ...maxScores } : null;
        },
        get turnCount() {
          return turnCountServer;
        },
        get turnState() {
          return turnState;
        },
        get result() {
          return serverResult;
        },
        get metering() {
          return metering;
        },
        get images() {
          return [...imagesByToken.values()];
        },
        get sawAction() {
          return sawAction;
        },

        [Symbol.asyncIterator]() {
          return {
            next(): Promise<IteratorResult<string>> {
              // Drain buffered deltas BEFORE surfacing done/error, so a fast turn
              // that finished before iteration began still yields all its text.
              if (queue.length > 0) {
                return Promise.resolve({ value: queue.shift()!, done: false });
              }
              if (streamError) return Promise.reject(streamError);
              if (streamDone) {
                return Promise.resolve({ value: undefined as any, done: true });
              }
              return new Promise((res, rej) => {
                waiting = { resolve: res, reject: rej };
              });
            },
          };
        },
      };

      const ws = new WebSocket(turnContext.socketUrl, { headers: turnContext.headers });

      let handshakeDone = false;
      let stopped = false;

      // Cancellation: when the caller's signal aborts (HTTP client disconnected),
      // cancel the in-flight M365 turn the same way the real UI does — send the
      // Stop frame, then let the server's type:3 ack close the socket (with a
      // hard-close fallback). Before the handshake completes there's nothing to
      // stop, so just close.
      const onAbort = () => {
        if (stopped) return;
        stopped = true;
        try {
          if (ws.readyState === WebSocket.OPEN && handshakeDone) {
            log.info("Aborted — sending Stop frame to cancel the turn");
            dumpFrame(requestId, { target: "stop", type: 1 }, "send");
            ws.send(STOP_FRAME);
            setTimeout(() => { try { ws.close(); } catch {} }, 2_000);
          } else {
            try { ws.close(); } catch {}
          }
        } catch {
          try { ws.close(); } catch {}
        }
      };
      if (signal) {
        if (signal.aborted) onAbort();
        else signal.addEventListener("abort", onAbort, { once: true });
      }
      const clearAbort = () => signal?.removeEventListener("abort", onAbort);

      ws.on("open", () => {
        log.debug("WS connected, sending handshake");
        ws.send(JSON.stringify({ protocol: "json", version: 1 }) + RS);
      });

      ws.on("message", (data: WebSocket.RawData) => {
        const raw = data.toString();
        log.debug("WS recv:", trunc(raw, 500));
        const frames = raw.split(RS).filter((f) => f.length > 0);

        for (const frame of frames) {
          let parsed: unknown;
          try {
            parsed = JSON.parse(frame);
          } catch {
            if (!handshakeDone) {
              handshakeDone = true;
              sendChat();
            }
            continue;
          }
          dumpFrame(requestId, parsed, "recv");

          if (!handshakeDone) {
            handshakeDone = true;
            const hs = SignalRHandshakeResponse.safeParse(parsed);
            if (hs.success && hs.data.error) {
              ws.close();
              reject(new Error(`Handshake error: ${hs.data.error}`));
              return;
            }
            sendChat();
            continue;
          }

          handleMsg(parsed);
        }
      });

      ws.on("error", (err: Error) => {
        const msg = err.message || "connection failed";
        log.error("WS error:", msg);
        if (!handshakeDone) {
          reject(new Error(`WebSocket error: ${msg}`));
        } else {
          onError(new Error(`WebSocket error: ${msg}`));
        }
      });

      ws.on("close", () => {
        clearAbort();
        log.info("WS closed, answer length:", answer.length);
        log.debug("Final response:", trunc(answer, 1000));
        onDone();
      });

      const sendChat = () => {
        const defaults = turnContext.payloadDefaults;
        const args = {
          ...defaults,
          source: "officeweb",
          clientCorrelationId: requestId,
          sessionId,
          optionsSets: [
            ...((defaults.optionsSets as string[] | undefined) ?? []),
            ...(supports("code_interpreter") ? CODE_INTERPRETER_OPTIONS_SETS : []),
            ...(wantImages ? IMAGE_GEN_OPTIONS_SETS : []),
            ...(process.env.M365_EXTRA_OPTIONSSETS ? process.env.M365_EXTRA_OPTIONSSETS.split(",").map((s) => s.trim()).filter(Boolean) : []),
          ],
          streamingMode: "ConciseWithPadding",
          spokenTextMode: "None",
          options: {},
          extraExtensionParameters: {},
          allowedMessageTypes: (defaults.allowedMessageTypes as string[] | undefined) ?? [
            ...ENTERPRISE_MESSAGE_TYPES,
            ...(wantImages ? ["GenerateGraphicArt"] : []),
            ...(nativeActions ? ACTION_ALLOWED_MESSAGE_TYPES : []),
          ],
          sliceIds: [] as string[],
          threadLevelGptId: agentId
            ? { id: agentId, source: "MOS3" }
            : {},
          traceId: requestId,
          isStartOfSession: isFirst,
          clientInfo: {
            clientPlatform: "mcmcopilot-web",
            clientAppName: "Office",
            clientEntrypoint: "mcmcopilot-officeweb",
            clientSessionId: sessionId,
            ...((defaults.clientInfo as Record<string, string> | undefined) ?? {}),
            clientAppType: "Web",
            deviceOS: (defaults.deviceOS as string | undefined) ?? "Linux",
            deviceType: "Desktop",
          },
          message: {
            author: "user",
            inputMethod: "Keyboard",
            text,
            entityAnnotationTypes: [
              "People",
              "File",
              "Event",
              "Email",
              "TeamsMessage",
            ],
            requestId,
            locationInfo: { timeZoneOffset: 1, timeZone: "Europe/Copenhagen" },
            locale: "en-gb",
            messageType: "Chat",
            experienceType: "Default",
            adaptiveCards: [] as any[],
            clientPreferences: {},
            ...(images.length ? { messageAnnotations: images } : {}),
          },
          ...(agentId
            ? {
                gpts: [{
                  id: agentId,
                  source: "MOS3",
                  version: "1.0.0",
                  clientOverrides: {
                    // The real capability channel (H8.2/H8.3/H8.7) — RegisteredPlugins,
                    // ScenarioModels, CodeInterpreter, … ride here per the decompile.
                    capabilities: nativeActions?.capabilities ?? [],
                    "deepResearchModels@odata.type": "Collection(String)",
                  },
                }],
              }
            : {}),
          // Inline agent definition — attach a custom action WITHOUT sideloading
          // through Teams/Graph (H-NATIVE-7). Best-effort schema from the decompile;
          // run with M365_DUMP_FRAMES=1 to see what the server accepts.
          ...(nativeActions?.gptDefinitions ? { gptDefinitions: nativeActions.gptDefinitions } : {}),
          // plugins[]: an explicit native-action list wins; otherwise Bing on the
          // plain (agent-less) path only, exactly as before.
          ...(nativeActions?.plugins
            ? { plugins: nativeActions.plugins }
            : (!agentId ? { plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }] } : {})),
          // Skill flags that gate the action/confirmation flow — the real client
          // sends these in the chat payload (bundle 8af68b68f4a2 destructure at
          // :9373-9378 → payload :11111). Without them the server has no reason to
          // run the confirmation-dialog / auto-invoke skills, so an attached action
          // never triggers. Only sent on the native-action path.
          ...(nativeActions
            ? {
                enableConfirmationDialogSkill: true,
                enableAgentAutoInvoke: true,
                enableMsgExtAuthSkill: true,
                enablePPCAuthSkill: true,
              }
            : {}),
          isSbsSupported: true,
          tone: (defaults.tone as string | undefined) ?? tone,
          renderReferencesBehindEOS: true,
          disconnectBehavior: "continue",
        };
        baseArgs = args;
        const chatMsg = {
          arguments: [args],
          invocationId: "0",
          target: "chat",
          type: 4,
        };

        const metrics = {
          arguments: [
            {
              Timestamps: {
                ConnectionStart: new Date().toISOString(),
                UserInputStart: new Date().toISOString(),
                ConnectionEstablished: new Date().toISOString(),
                UserInputSubmit: new Date().toISOString(),
              },
            },
          ],
          target: "Metrics",
          type: 1,
        };

        const payload = JSON.stringify(chatMsg) + RS + JSON.stringify(metrics) + RS;
        log.debug("WS send:", trunc(payload, 500));
        dumpFrame(requestId, chatMsg, "send");
        dumpFrame(requestId, metrics, "send");
        ws.send(payload);
        resolve(stream);
      };

      // When the model triggers a native custom action, the server sends a
      // confirmation TRIGGER (an adaptive card) and waits for us to approve it on
      // THIS socket. We reuse the exact chat envelope (baseArgs) with a
      // ResumeInvokeAction message, so the server-side orchestrator makes the real
      // outbound call (H-NATIVE-6). Returns true if it sent a resume (⇒ keep socket
      // open for the result). No-op unless native actions are enabled.
      const maybeResumeAction = (m: unknown): boolean => {
        if (!nativeActions || actionResumed || !baseArgs) return false;
        const conf = parseActionConfirmation(m as any);
        if (!conf) return false;
        sawAction = true;
        if (!shouldAutoConfirm(conf, { autoConfirmAll: nativeActions.autoConfirmAll })) {
          log.info(`Native action ${conf.actionId} is consequential; autoConfirmAll off — not resuming`);
          return false;
        }
        actionResumed = true;
        const resumeArgs = { ...baseArgs, message: buildResumeInvokeAction(conf) };
        const frame = { arguments: [resumeArgs], invocationId: "0", target: "chat", type: 4 };
        dumpFrame(requestId, frame, "send");
        log.info(`Native action: auto-confirming actionId=${conf.actionId} via ResumeInvokeAction`);
        try { ws.send(JSON.stringify(frame) + RS); } catch (e) { log.error("resume send failed:", (e as Error).message); }
        return true;
      };

      function handleMsg(raw: unknown) {
        const base = raw as { type?: number; target?: string; arguments?: unknown[] };

        if (base.type === 6) {
          ws.send(JSON.stringify({ type: 6 }) + RS);
          return;
        }

        if (base.type === 7) {
          const frame = CloseFrame.safeParse(raw);
          if (frame.success && frame.data.error) {
            onError(new Error(`Server close: ${frame.data.error}`));
          }
          ws.close();
          return;
        }

        if (base.type === 3) {
          const frame = CompletionFrame.safeParse(raw);
          if (frame.success && frame.data.error) {
            onError(new Error(`Completion error: ${frame.data.error}`));
          }
          ws.close();
          return;
        }

        if (base.type === 2) {
          // Stream item — the FINAL state of the conversation, with authoritative
          // throttle/turnCount/scores. Mine it before closing.
          const item = (raw as { item?: { messages?: any[]; throttling?: any; turnState?: string; result?: { value?: unknown; errorCode?: unknown; message?: unknown } } }).item;
          if (item) {
            if (typeof item.result?.value === "string") {
              serverResult = {
                value: item.result.value,
                ...(typeof item.result.errorCode === "string" ? { errorCode: item.result.errorCode } : {}),
                ...(typeof item.result.message === "string" ? { message: item.result.message } : {}),
              };
              if (item.result.value !== "Success") log.info(`Turn result: ${item.result.value}${serverResult.errorCode ? ` (${serverResult.errorCode})` : ""}${accountTag}`);
              if (item.result.value === "Throttled") recordThrottle(Date.now(), accountId);
            }
            if (item.turnState) turnState = item.turnState;
            if (item.throttling) {
              throttleInfo = { current: item.throttling.numUserMessagesInConversation, max: item.throttling.maxNumUserMessagesInConversation };
              metering = parseMetering(item.throttling.metering);
              for (const [budget, { label }] of Object.entries(METERED_BUDGETS)) {
                const left = meteredAllowance(metering, budget as MeteredBudget);
                if (left) log.info(`${label} priority access left: ${left.daily ?? "?"} today, ${left.weekly ?? "?"} this week`);
              }
            }
            let resumedHere = false;
            for (const m of item.messages ?? []) {
              if (m.author !== "bot") continue;
              captureImages(m);
              if (m.contentOrigin) contentOrigin = m.contentOrigin;
              if (m.messageType) messageType = m.messageType;
              if (m.messageId) messageId = m.messageId;
              if (typeof m.turnCount === "number") turnCountServer = m.turnCount;
              if (Array.isArray(m.scores)) {
                for (const s of m.scores) {
                  if (typeof s?.component !== "string" || typeof s?.score !== "number") continue;
                  if (!(s.component in maxScores) || s.score > maxScores[s.component]) {
                    maxScores[s.component] = s.score;
                  }
                }
              }
              if (maybeResumeAction(m)) resumedHere = true;
            }
            // If the model just triggered a custom action, this "final" item isn't
            // final — we sent a ResumeInvokeAction and must keep the socket open to
            // receive the action's result (H-NATIVE-6).
            if (resumedHere) return;
          }
          ws.close();
          return;
        }

        if (base.type === 1 && base.target === "update" && Array.isArray(base.arguments)) {
          for (const arg of base.arguments) {
            const delta = DeltaUpdate.safeParse(arg);
            if (delta.success) {
              composer.delta(delta.data.writeAtCursor);
              advance(composer.text);
              continue;
            }

            const msgUpdate = MessageUpdate.safeParse(arg);
            if (msgUpdate.success) {
              // A cursor names the message that the following deltas extend.
              const cursorId = cursorMessageId(msgUpdate.data.cursor?.j);
              if (cursorId) composer.cursor(cursorId);
              for (const m of msgUpdate.data.messages) {
                // Capture diagnostic meta from every bot message — including the
                // control-typed ones — so callers can tell apart `DeepLeo` from
                // `3PDeclarativeAgent` and surface `Disengaged` cleanly.
                if (m.author === "bot") {
                  captureImages(m);
                  if (m.contentOrigin) contentOrigin = m.contentOrigin;
                  if (m.messageType) messageType = m.messageType;
                  if (m.messageId) messageId = m.messageId;
                  if (m.scores) {
                    for (const s of m.scores) {
                      if (!(s.component in maxScores) || s.score > maxScores[s.component]) {
                        maxScores[s.component] = s.score;
                      }
                    }
                  }
                  if (typeof m.turnCount === "number") turnCountServer = m.turnCount;
                  if (m.turnState) turnState = m.turnState;
                  // A confirmation TRIGGER for a native action arrives here as a
                  // messageType'd bot frame (which the plain-chat path drops). Detect
                  // it and auto-approve on the same socket (H-NATIVE-6).
                  maybeResumeAction(m);
                }
                if (m.author === "bot" && m.text && !m.messageType) {
                  composer.snapshot(m.messageId, m.text);
                  advance(composer.text);
                }
              }
              continue;
            }

            const throttle = ThrottlingUpdate.safeParse(arg);
            if (throttle.success) {
              const t = throttle.data.throttling;
              throttleInfo = { current: t.numUserMessagesInConversation, max: t.maxNumUserMessagesInConversation };
              log.info(`Throttle: ${t.numUserMessagesInConversation}/${t.maxNumUserMessagesInConversation} messages`);
            }
          }
        }
      }
    });
  }
}
