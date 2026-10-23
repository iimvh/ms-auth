import { PersonalAuthError, personalLoginCommand } from "../adapters/personal/personal-auth.js";
import type { PersonalCapture, PersonalRecord } from "../adapters/personal/personal-session.js";
import { PersonalStore } from "../adapters/personal/personal-session.js";
import { loginCommandFor, resolveAccount } from "../account-registry.js";
import { ACTION_ALLOWED_MESSAGE_TYPES } from "../tools/native-actions.js";

/**
 * Build the authenticated Chathub URL. Kept as a pure helper so privacy-critical
 * query flags can be verified without opening a network connection.
 */
export function buildCopilotWebSocketUrl(
  objectId: string,
  tenantId: string,
  params: URLSearchParams,
): string {
  return `wss://substrate.office.com/m365Copilot/Chathub/${objectId}@${tenantId}?${params}`;
}

/** Per-turn query params for the personal path: the captured connection
 *  values (scenario, variants, licenseType, …), not the enterprise scenario
 *  table — except `M365_SCENARIO`, which still wins if set (operator
 *  override, spec §6.4). `disableMemory=1` is still added by the same place
 *  that adds it today when `temporaryChat`. */
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

/** The Chathub URL for the captured session: same host as the enterprise
 *  path, but the connection identity comes from the capture, not a JWT's
 *  oid@tid claim. */
export function buildPersonalChathubUrl(
  capture: Pick<PersonalCapture, "chathubPath">,
  params: URLSearchParams,
): string {
  return `wss://substrate.office.com/m365Copilot/Chathub/${capture.chathubPath}?${params}`;
}

/** The record backing a personal turn. The pre-flight (ensurePersonalCredentials)
 *  guarantees one exists; if it still does not — or is corrupt, or has expired —
 *  fail with an error that names WHICH account and why, not a mystery 401 and
 *  not a flat `session_dead` (§7.5). With several accounts, "which one" is the
 *  first thing the operator needs.
 *
 *  `accountId` pins the account: a thread passes the one it was opened on, so
 *  its record cannot come from a different account than the token the caller
 *  resolved from the same id. Omitted = the active account. */
export function readPersonalRecord(accountId?: string | null): PersonalRecord {
  const account = accountId === undefined ? resolveAccount() : resolveAccount({ id: accountId });
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

/**
 * Headers for the Chathub WebSocket upgrade.
 *
 * `X-RoutingParameter-SessionKey` pins every turn of a conversation to one
 * backend. Each turn is a new connection, and without the key the Substrate
 * front end spreads them over pods in two regions (`x-calculatedbetarget`:
 * 4–6 distinct in 6 handshakes on one ConversationId, Switzerland North and
 * Sweden Central), which keep separate copies of the conversation. The model
 * then answers from whichever copy the turn landed on: in real pi runs 53% of
 * turns ran on an older copy, missing the model's own earlier calls and their
 * results (docs/hypotheses.md F85). With the key: 1 backend in 6. The query-string form
 * doesn't pin. M365_NO_SESSION_ROUTING=1 leaves it out.
 */
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

/** Headers for the personal (consumer) Chathub socket. The consumer substrate
 *  cancels the handshake ("Handshake was canceled.") unless the upgrade looks
 *  like the real Windows Chrome client: the browser UA plus the cache and
 *  encoding headers it always sends. Cookie/Turnstile proof are NOT needed once
 *  this matches (F103, live replay). Do NOT hand-set Sec-WebSocket-Extensions:
 *  the websocket lib negotiates permessage-deflate itself, and a manual copy
 *  makes it skip decompression — the stream then dies as "Unexpected end of
 *  JSON input" on the first compressed frame (F105). */
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
    // Image generation (§14): the server only SENDS the GraphicArt
    // frame if the client declares it can handle it — same
    // declare-to-receive rule as native actions.
    ...(wantImages ? ["GenerateGraphicArt"] : []),
    // Native custom-action vocabulary (H-NATIVE-6): the server only
    // SENDS these trigger frames if the client says it can handle them.
    ...(nativeActions ? ACTION_ALLOWED_MESSAGE_TYPES : []),
  ];
}

export function personalPayloadClientInfo(): Record<string, string> {
  return {
    // The consumer client's extra descriptors (F108) — the
    // enterprise backend ignores extras, the consumer one is
    // stricter about which surface is speaking.
    ProductCategory: "Chat",
    productEntryPoint: "ChatPanel",
    clientPlatformVersion: "10",
  };
}

export const PERSONAL_PAYLOAD_DEVICE_OS = "Windows";

// Consumer (personal) accounts serve the default Copilot model only:
// asking for a Claude_Sonnet tone answers InternalError on an
// otherwise identical turn (F108, live A/B). The real client sends
// "Magic", capitalised.
export const PERSONAL_PAYLOAD_TONE = "Magic";
