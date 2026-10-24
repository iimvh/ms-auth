import { getToken } from "../adapters/msal/auth.js";
import { getPersonalToken } from "../adapters/personal/personal-auth.js";
import { isPersonalMode, shouldSendCodeInterpreter } from "../adapters/personal/personal-mode.js";
import type { PersonalRecord } from "../adapters/personal/personal-session.js";
import { resolveAccount } from "../account-registry.js";
import { decodeJwt, getScenarioForModel } from "../routing/copilot.js";
import type { TurnContext } from "../turn-context.js";
import {
  buildCopilotWebSocketHeaders,
  buildCopilotWebSocketUrl,
  buildPersonalChathubUrl,
  buildPersonalWebSocketHeaders,
  personalPayloadAllowedMessageTypes,
  personalPayloadClientInfo,
  personalPayloadOptionsSets,
  personalTurnParams,
  readPersonalRecord,
  PERSONAL_PAYLOAD_DEVICE_OS,
  PERSONAL_PAYLOAD_TONE,
} from "../personal-chathub/chathub.js";

const ENTERPRISE_VARIANTS = [
  "EnableMcpServerWidgets",
  "feature.EnableMcpServerWidgets",
  "feature.EnableLuForChatCIQ",
  "feature.enableChatCIQPlugin",
  "EnableRequestPlugins",
  "feature.EnableSensitivityLabels",
  "EnableUnsupportedUrlDetector",
  "feature.IsCustomEngineCopilotEnabled",
  "feature.bizchatfluxv3",
  "feature.enablechatpages",
  "feature.enableCodeCanvas",
  "feature.turnOnWorkTabRecommendation",
  "turnOffWorkTabUpsellFromClient",
  "feature.turnOnDARecommendation",
  "feature.IsStreamingModeInChatRequestEnabled",
  "IncludeSourceAttributionsConcise",
  "SkipPublishEmptyMessage",
  "feature.EnableDeduplicatingSourceAttributions",
  "Enable3PActionProgressMessages",
  "feature.enableClientWebRtc",
  "feature.EnableMeetingRecapOfSeriesMeetingWithCiq",
  "feature.EnableReferencesListCompleteSignal",
  "feature.StorageMessageSplitDisabled",
  "feature.EnableCuaTakeControlApi",
  "feature.cwcallowedos",
  "feature.disabledisallowedmsgs",
  "feature.enableCitationsForSynthesisData",
  "feature.enableGenerateGraphicArtOptionsSet",
  "cdximagen",
  "feature.EnableUpdatedUXForConfirmationDialog",
  "feature.EnableClientFileURLSupportForOfficeWebPaidCopilot",
  "feature.EnableDesignEditorImageGrounding",
  "feature.EnableDesignerEditor",
  "feature.OfficeWebToHelix",
  "feature.OfficeDesktopToHelix",
  "feature.M365TeamsHubToHelix",
  "feature.OwaHubToHelix",
  "feature.MonarchHubToHelix",
  "feature.Win32OutlookHubToHelix",
  "feature.MacOutlookHubToHelix",
  "Agt_bizchat_enableGpt5ForHelix",
].join(",");

export interface CredentialSelection {
  accountId: string | null;
  personal: boolean;
  token: string;
  record: PersonalRecord | null;
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

export async function selectCredential(accountId?: string | null): Promise<CredentialSelection> {
  if (!isPersonalMode()) {
    return { accountId: null, personal: false, token: await getToken(), record: null };
  }
  const account = accountId === undefined ? resolveAccount() : resolveAccount({ id: accountId });
  const record = readPersonalRecord(account);
  const token = await getPersonalToken({ accountId: account.id, record });
  return { accountId: account.id, personal: true, token, record };
}

export function turnContextFor(selection: CredentialSelection, plan: TurnPlan): TurnContext {
  return selection.personal
    ? personalTurnContext(selection, plan)
    : enterpriseTurnContext(selection, plan);
}

export function agentAllowed(selection: CredentialSelection, requested: boolean): boolean {
  return requested && !selection.personal;
}

export function imagesAllowed(selection: CredentialSelection, agentId: string | undefined): boolean {
  return !selection.personal && !agentId && !process.env.M365_NO_IMAGE_GEN;
}

function personalTurnContext(selection: CredentialSelection, plan: TurnPlan): TurnContext {
  const record = selection.record!;
  const params = personalTurnParams(record, {
    requestId: plan.requestId,
    sessionId: plan.sessionId,
    conversationId: plan.conversationId,
  });
  if (plan.temporaryChat) params.set("disableMemory", "1");
  return {
    token: selection.token,
    accountId: selection.accountId,
    personal: true,
    socketUrl: buildPersonalChathubUrl(record, params),
    headers: buildPersonalWebSocketHeaders(plan.conversationId),
    payloadDefaults: {
      optionsSets: personalPayloadOptionsSets(record),
      allowedMessageTypes: personalPayloadAllowedMessageTypes(record, plan.images, plan.nativeActions),
      clientInfo: personalPayloadClientInfo(),
      deviceOS: PERSONAL_PAYLOAD_DEVICE_OS,
      tone: PERSONAL_PAYLOAD_TONE,
    },
    capabilities: shouldSendCodeInterpreter(undefined, true) ? ["code_interpreter"] : [],
    record,
  };
}

function enterpriseTurnContext(selection: CredentialSelection, plan: TurnPlan): TurnContext {
  const { scenario, licenseType } = getScenarioForModel(plan.model);
  const params = new URLSearchParams({
    chatsessionid: plan.requestId,
    clientrequestid: plan.requestId,
    "X-SessionId": plan.sessionId,
    ConversationId: plan.conversationId,
    access_token: selection.token,
    variants: ENTERPRISE_VARIANTS,
    source: '"officeweb"',
    product: "Office",
    agentHost: "Bizchat.FullScreen",
    licenseType,
    agent: "web",
    scenario,
  });
  if (plan.temporaryChat) params.set("disableMemory", "1");
  const claims = claimsOrEmpty(selection.token);
  return {
    token: selection.token,
    accountId: null,
    personal: false,
    socketUrl: buildCopilotWebSocketUrl(claims.oid, claims.tid, params),
    headers: buildCopilotWebSocketHeaders(plan.conversationId),
    payloadDefaults: {},
    capabilities: shouldSendCodeInterpreter(plan.agentId, false) ? ["code_interpreter"] : [],
    record: null,
  };
}

function claimsOrEmpty(token: string): { oid: string; tid: string } {
  try {
    const claims = decodeJwt(token);
    return { oid: claims.oid, tid: claims.tid };
  } catch {
    return { oid: "", tid: "" };
  }
}
