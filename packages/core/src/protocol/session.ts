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

const STOP_FRAME = JSON.stringify({ arguments: [{}], invocationId: "1", target: "stop", type: 1 }) + RS;

const CODE_INTERPRETER_OPTIONS_SETS = [
  "cwc_code_interpreter",
  "cwc_code_interpreter_amsfix",
  "cwc_code_interpreter_citation_fix",
  "code_interpreter_interactive_charts",
  "code_interpreter_matplotlib_patching",
];

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

// M365_DUMP_FRAMES=1 appends every frame to <FRAME_DIR>/<requestId>.ndjson.
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
  }
}

export function foldStreamText(
  answer: string,
  next: string,
): { answer: string; emit: string | null } {
  if (next.length <= answer.length) return { answer, emit: null };
  if (next.startsWith(answer)) return { answer: next, emit: next.slice(answer.length) };
  return { answer: next, emit: null };
}

export function cursorMessageId(j: string | undefined): string | null {
  const m = j?.match(/^\$\['([^']+)'\]/);
  return m ? m[1] : null;
}

export class TurnTextComposer {
  private texts = new Map<string, string>();
  private target: string | null = null;

  cursor(id: string): void {
    this.target = id;
  }

  snapshot(id: string | undefined, text: string): void {
    const key = id ?? this.target ?? "";
    if (this.target === null) this.target = key;
    this.texts.set(key, foldStreamText(this.texts.get(key) ?? "", text).answer);
  }

  delta(text: string): void {
    const key = this.target ?? "";
    this.texts.set(key, (this.texts.get(key) ?? "") + text);
  }

  get text(): string {
    return [...this.texts.values()].filter((t) => t.length > 0).join("\n\n");
  }
}

export interface NativeActionConfig {
  gptDefinitions?: unknown[];
  capabilities?: unknown[];
  plugins?: unknown[];
  autoConfirmAll?: boolean;
}

export interface ChatTurnOptions {
  generateImages?: boolean;
  images?: ImageAnnotation[];
}

export interface CopilotSessionOptions {
  agentId?: string;
  sessionId?: string;
  conversationId?: string;
  temporaryChat?: boolean;
  nativeActions?: NativeActionConfig;
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

  get turnCount(): number {
    return this._turnCount;
  }

  get account(): string | null {
    return this.accountId;
  }

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
      const maxScores: Record<string, number> = {};
      let turnCountServer: number | null = null;
      let turnState: string | null = null;
      let serverResult: { value: string; errorCode?: string; message?: string } | null = null;
      let metering: Record<string, number> | null = null;
      const imagesByToken = new Map<string, CapturedImage>();
      let baseArgs: Record<string, unknown> | null = null;
      let sawAction = false;
      let actionResumed = false;

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
                    capabilities: nativeActions?.capabilities ?? [],
                    "deepResearchModels@odata.type": "Collection(String)",
                  },
                }],
              }
            : {}),
          ...(nativeActions?.gptDefinitions ? { gptDefinitions: nativeActions.gptDefinitions } : {}),
          ...(nativeActions?.plugins
            ? { plugins: nativeActions.plugins }
            : (!agentId ? { plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }] } : {})),
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
              const cursorId = cursorMessageId(msgUpdate.data.cursor?.j);
              if (cursorId) composer.cursor(cursorId);
              for (const m of msgUpdate.data.messages) {
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
