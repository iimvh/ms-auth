import WebSocket from "ws";
import type { PersonalCapture } from "../adapters/personal/personal-session.js";
import { createLogger } from "../io/log.js";
import { buildPersonalChathubUrl, buildPersonalWebSocketHeaders, personalTurnParams } from "./chathub.js";

const RS = "\x1E";
const log = createLogger("session");

export async function probePersonalCapture(
  capture: PersonalCapture,
  models: string[],
): Promise<{ model: string; ok: boolean; at: number }> {
  for (const model of models) {
    log.info(`personal capability probe: trying ${model}`);
    const ok = await probeModel(capture, model);
    if (ok) return { model, ok: true, at: Date.now() };
  }
  return { model: models[0]!, ok: false, at: Date.now() };
}

const PROBE_TIMEOUT_MS = 30_000;

async function probeModel(capture: PersonalCapture, model: string): Promise<boolean> {
  const requestId = crypto.randomUUID();
  const params = personalTurnParams(capture, {
    requestId,
    sessionId: crypto.randomUUID(),
    conversationId: crypto.randomUUID(),
  });
  const wsUrl = buildPersonalChathubUrl(capture, params);
  return new Promise((resolve) => {
    const ws = new WebSocket(wsUrl, { headers: buildPersonalWebSocketHeaders(params.get("ConversationId")!) });
    const done = (ok: boolean) => {
      clearTimeout(timer);
      if (ws.readyState !== WebSocket.CLOSED) ws.close();
      resolve(ok);
    };
    const timer = setTimeout(() => done(false), PROBE_TIMEOUT_MS);
    let handshakeDone = false;

    ws.on("open", () => ws.send(JSON.stringify({ protocol: "json", version: 1 }) + RS));
    ws.on("error", (err) => {
      log.debug(`probe error for ${model}: ${err.message}`);
      done(false);
    });
    ws.on("close", () => done(false));
    ws.on("message", (data: WebSocket.RawData) => {
      for (const line of data.toString().split(RS)) {
        if (!line) continue;
        let parsed: any;
        try {
          parsed = JSON.parse(line);
        } catch {
          continue;
        }
        if (!handshakeDone) {
          handshakeDone = true;
          ws.send(
            JSON.stringify({
              arguments: [{
                source: capture.source,
                clientCorrelationId: requestId,
                sessionId: params.get("X-SessionId")!,
                optionsSets: capture.optionsSets,
                streamingMode: "ConciseWithPadding",
                spokenTextMode: "None",
                options: {},
                extraExtensionParameters: {},
                allowedMessageTypes: capture.allowedMessageTypes,
                sliceIds: [] as string[],
                threadLevelGptId: {},
                traceId: requestId,
                isStartOfSession: true,
                clientInfo: {
                  clientPlatform: "mcmcopilot-web",
                  clientAppName: "Office",
                  clientEntrypoint: "mcmcopilot-officeweb",
                  clientSessionId: params.get("X-SessionId")!,
                  clientAppType: "Web",
                  deviceOS: "Linux",
                  deviceType: "Desktop",
                },
                message: {
                  author: "user",
                  inputMethod: "Keyboard",
                  text: "reply with the single word ok",
                  entityType: 0,
                },
                tone: "Magic",
              }],
              invocationId: "0",
              target: "chat",
              type: 4,
            }) + RS,
          );
          continue;
        }
        const messages = parsed?.item?.messages ?? parsed?.arguments?.[0]?.messages;
        if (Array.isArray(messages)) {
          const botText = messages.some(
            (m: any) =>
              m.author === "bot" &&
              typeof m.text === "string" &&
              m.text.trim() !== "",
          );
          if (botText) done(true);
        }
      }
    });
  });
}
