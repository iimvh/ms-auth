import "./personal-wiring.js";

export {
  getToken,
  getTokenSilent,
  getTokenForScope,
  getImageArtifactToken,
  loginAutomated,
  loadSecrets,
  forceReauth,
} from "./adapters/msal/auth.js";

export {
  generateImage,
  fetchImageBytes,
  buildImagePrompt,
  classifyImageFailure,
  ImageGenerationError,
  type GeneratedImage,
  type GenerateImageOptions,
  type ImageOrientation,
  type ImageStyle,
  type ImageGenFailureReason,
} from "./io/image.js";

export { uploadImage, imageUrls, type ImageAnnotation } from "./io/image-input.js";

export {
  AccountError,
  ACCOUNT_CDP_PORT,
  ACCOUNT_ID_RE,
  ACCOUNT_PROFILE_DIR,
  ACCOUNT_SESSION_FILE,
  accountIds,
  accountKeyOf,
  accountRegistrySource,
  accountsDir,
  accountsRoot,
  accountViews,
  activateAccount,
  addAccount,
  allocateAccount,
  chathubOid,
  chromeCommandFor,
  clearMockRegistry,
  configDir,
  createMockRegistry,
  createMockRegistryAt,
  expandHome,
  isSafeAccountId,
  loginCommandFor,
  migrateRegistry,
  readRegistry,
  registryProblem,
  registryStatus,
  removeAccount,
  renameAccount,
  resolveAccount,
  useMockRegistry,
  writeRegistry,
  type AccountEntry,
  type AccountId,
  type AccountPaths,
  type AccountRegistry,
  type AccountView,
  type NewAccount,
  type RegistryRead,
  type RegistryStatus,
  type ResolvedAccount,
  type ResolveAccountOptions,
} from "./account-registry.js";

export {
  noteRequestOutcome,
  awaitDegradationBackoff,
  isDegradationBackoff,
  forgetAccountBackoff,
  createBackoffController,
  type BackoffController,
  type BackoffOptions,
} from "./budget/auth-recovery.js";

export { getOrCreateAgent, getEnvironmentUrl, environmentUrlFromName } from "./io/agent.js";

export {
  decodeJwt,
  getToneForModel,
  getAvailableModels,
  getScenarioForTone,
  getScenarioForModel,
  isSonnet5Model,
  isSonnet55Model,
  isIncludedOpusModel,
  toneUsesToolAgent,
  toolRequestUsesAgent,
  isPremiumOnlyModel,
  requestUsesAgent,
  meteredBudgetOf,
  priorityAccessFallbackModel,
  PRIORITY_ACCESS_FALLBACK_ENV,
  noteAgentRouteDead,
  noteAgentRouteAlive,
  isAgentRouteAlive,
  resetAgentRoutes,
  AGENTLESS_TOOL_TONES,
  AGENT_CLAUDE_TONES,
  PREMIUM_ONLY_AGENT_TONES,
  PAID_SCENARIO_TONES,
  PAID_SCENARIO_MODELS,
  INCLUDED_SCENARIO_MODELS,
  type ScenarioRouting,
  type CopilotStream,
  type CapturedImage,
} from "./routing/copilot.js";

export {
  parsePriorityAccessExhaustion,
  couldBePriorityAccessPrefix,
  secondsUntilReset,
  parseMetering,
  meteredAllowance,
  METERED_BUDGETS,
  priorityAccessExhaustionOf,
  notePriorityAccessExhausted,
  activePriorityAccessExhaustion,
  resetPriorityAccessState,
  type MeteredBudget,
  type PriorityAccessExhaustion,
  type PriorityAccessWindow,
} from "./budget/priority-access.js";

export {
  currentFramingVariant,
  defaultFramingForTone,
  defaultFramingForModel,
  transcriptStyleForVariant,
  FRAMING_VARIANT_NAMES,
  longHeredocAsWrite,
  type TranscriptStyle,
} from "./framing/fenced.js";

export {
  CopilotSession,
  type CopilotSessionOptions,
  type ChatTurnOptions,
  type NativeActionConfig,
} from "./protocol/session.js";

export {
  personalTurnParams,
  buildPersonalChathubUrl,
  readPersonalRecord,
  buildPersonalWebSocketHeaders,
} from "./personal-chathub/chathub.js";

export { probePersonalCapture } from "./personal-chathub/probe.js";

export type { TurnContext, CredentialSource, SessionStore } from "./turn-context.js";

export {
  parseActionConfirmation,
  buildResumeInvokeAction,
  shouldAutoConfirm,
  buildNativeActionPrompt,
  NATIVE_ACTION_INSTRUCTIONS,
  ACTION_ALLOWED_MESSAGE_TYPES,
  ACTION_CONFIRM_MESSAGE_TYPES,
  type ActionConfirmation,
} from "./tools/native-actions.js";

export {
  ModelSession,
  type ModelSessionOptions,
} from "./app/model.js";

export { isPersonalMode, resolveAuthMode, personalDefaultModel, personalModeDefaults, shouldSendCodeInterpreter, type AuthMode } from "./adapters/personal/personal-mode.js";
export { getPersonalToken, ensurePersonalCredentials, personalLoginCommand, PersonalAuthError, PERSONAL_AUTH_ERROR_CODES, type PersonalAuthErrorCode } from "./adapters/personal/personal-auth.js";
export { PERSONAL_MINT_RATE_LIMITS, PERSONAL_MINT_RATE_LIMIT_ENV, MINT_HOUR_MS, PersonalMintRateLimitOverrideSchema, personalMintRateLimits, type PersonalMintRateLimit, type PersonalMintRateLimitOverride } from "./adapters/personal/personal-mint-limits.js";
export { PersonalStore, personalSessionFile, personalProfileDir, PersonalCaptureSchema, PersonalRecordSchema, PersonalRecordProblemSchema, type PersonalCapture, type PersonalRecord, type PersonalRecordProblem, type PersonalRecordState } from "./adapters/personal/personal-session.js";
export { capturePersonalSession, parsePersonalCapture, isLoginUrl, frameText, hasChatFrame, PersonalCaptureError, decodeJwtExpiresAt, type PersonalCaptureErrorReason } from "./adapters/personal/personal-capture.js";
export { harvestPersonalToken, tokenFromSocketUrl, HarvestError, type HarvestErrorReason, type HarvestResult, type HarvestedPage } from "./adapters/personal/personal-harvest.js";
export { attachChatPage, personalCdpUrl } from "./adapters/personal/personal-harvest-cdp.js";
export { personalHealth, probeCdpReachable, type PersonalHealth, type PersonalHealthDeps } from "./adapters/personal/personal-health.js";
export { driveSignInForm } from "./adapters/msal/auth.js";

export {
  formatMessages,
  formatToolDefinitions,
  formatToolChoiceInstruction,
  getMessageContent,
  parseToolCalls,
  looksLikeConfabulation,
  looksLikeHallucinatedCompletion,
  looksLikeRemoteArtifactCompletion,
  truncateAtFabricatedToolResponse,
  fabricatedToolResponseAt,
  textAfterFirstToolCall,
  isProseDocument,
  formatToolResponse,
  type Message,
  type ToolDef,
  type ToolFunction,
  type ToolChoice,
  type ParsedToolCall,
  type ParseResult,
} from "./tools/tools.js";

export { createLogger, trunc, LOG_PATH, FRAME_DIR, fileLoggingEnabled } from "./io/log.js";
export {
  budgetConfig,
  parseEvents,
  findLogs,
  readLogEvents,
  replay,
  decide,
  budgetStatus,
  proxyBudgetStatus,
  awaitTurnBudget,
  recordTurn,
  recordThrottle,
  resetTurnBudget,
  forgetAccountBudget,
  TurnBudgetHold,
  type BudgetConfig,
  type BudgetEvent,
  type BudgetState,
  type BudgetDecision,
  type AwaitTurnBudgetOptions,
} from "./budget/turn-budget.js";