import { PersonalAuthError, personalLoginCommand } from "../adapters/personal/personal-auth.js";
import type { PersonalCapture, PersonalRecord } from "../adapters/personal/personal-session.js";
import { PersonalStore } from "../adapters/personal/personal-session.js";
import type { ResolvedAccount } from "../account-registry.js";
import { loginCommandFor } from "../account-registry.js";
import { ACTION_ALLOWED_MESSAGE_TYPES } from "../tools/native-actions.js";

export function buildCopilotWebSocketUrl(
  objectId: string,
  tenantId: string,
  params: URLSearchParams,
): string {
  return `wss://substrate.office.com/m365Copilot/Chathub/${objectId}@${tenantId}?${params}`;
}

export function personalTurnParams(
  capture: PersonalCapture,
  ids: { requestId: string; sessionId: string; conversationId: string },
): URLSearchParams {
  return new URLSearchParams({
    chatsessionid: ids.requestId,
    clientrequestid: ids.requestId,
    "X-SessionId": ids.sessionId,
    ConversationId: ids.conversationId,
    access_token: capture.accessToken,
    variants: capture.variants,
    source: capture.source,
    agentHost: capture.agentHost,
    licenseType: capture.licenseType,
    agent: "web",
    scenario: process.env.M365_SCENARIO ?? capture.scenario,
    ...(capture.product === null ? {} : { product: capture.product }),
  });}

export function buildPersonalChathubUrl(
  capture: Pick<PersonalCapture, "chathubPath">,
  params: URLSearchParams,
): string {
  return `wss://substrate.office.com/m365Copilot/Chathub/${capture.chathubPath}?${params}`;
}

export function readPersonalRecord(account: ResolvedAccount): PersonalRecord {
  const whose = account.id ? `account "${account.id}"` : "the personal account";
  const login = account.id ? loginCommandFor(account.id) : personalLoginCommand();
  const store = new PersonalStore(account.sessionFile);
  const state = store.inspect();
  if (!state.ok) {
    const why =
      state.problem === "absent"
        ? `no session record at ${account.sessionFile}`
        : `its session record is unusable (${state.problem}: ${state.detail}) — delete ${account.sessionFile} (or pnpm m365-personal logout${account.id ? ` --account ${account.id}` : ""}) and run: ${login}`;
    throw new PersonalAuthError("session_dead", `M365 personal session expired for ${whose} — ${why}${state.problem === "absent" ? `; run: ${login}` : ""}`);
  }
  const record = state.record;
  if (!store.isFresh(record)) {
    const expired = new Date(record.expiresAt).toISOString();
    throw new PersonalAuthError(
      "session_dead",
      `M365 personal session expired for ${whose} — the record expired at ${expired}; run: ${login} (or wait for the next request to re-mint it from that account's Chrome)`,
    );
  }
  return record;
}

export function buildCopilotWebSocketHeaders(conversationId: string): Record<string, string> {
  return {
    "Origin": "https://m365.cloud.microsoft",
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64; rv:148.0) Gecko/20100101 Firefox/148.0",
    "Accept-Language": "en-US,en;q=0.9",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
    ...(process.env.M365_NO_SESSION_ROUTING ? {} : { "X-RoutingParameter-SessionKey": conversationId }),
  };
}

export function buildPersonalWebSocketHeaders(conversationId: string): Record<string, string> {
  return {
    Origin: "https://m365.cloud.microsoft",
    "User-Agent": process.env.M365_PERSONAL_UA ??
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Cache-Control": "no-cache",
    Pragma: "no-cache",
    ...(process.env.M365_NO_SESSION_ROUTING ? {} : { "X-RoutingParameter-SessionKey": conversationId }),
  };
}

export function personalPayloadOptionsSets(capture: PersonalCapture): string[] {
  return capture.optionsSets;
}

export function personalPayloadAllowedMessageTypes(
  capture: PersonalCapture,
  wantImages: boolean,
  nativeActions: boolean,
): string[] {
  return [
    ...capture.allowedMessageTypes,
    ...(wantImages ? ["GenerateGraphicArt"] : []),
    ...(nativeActions ? ACTION_ALLOWED_MESSAGE_TYPES : []),
  ];
}

export function personalPayloadClientInfo(): Record<string, string> {
  return {
    ProductCategory: "Chat",
    productEntryPoint: "ChatPanel",
    clientPlatformVersion: "10",
  };
}

export const PERSONAL_PAYLOAD_DEVICE_OS = "Windows";

export const PERSONAL_PAYLOAD_TONE = "Magic";
