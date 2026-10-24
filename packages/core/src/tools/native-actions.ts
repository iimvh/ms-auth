
export const NATIVE_ACTION_INSTRUCTIONS = `You are an automated agent with real, executable actions. When a request needs one, CALL the action — the runtime runs it against a live system and returns the real result. Rules:
- Never invent, guess, or approximate a value that an action returns. If an action can get it, call the action and report EXACTLY what it returns.
- Do not describe what you "would" do or ask the user to run anything. Call the action.
- After an action returns, answer using its real result — no preamble, no sign-off.

Available actions:
{{ACTIONS}}`;

export function buildNativeActionPrompt(actions: Array<{ name: string; description?: string }>): string {
  const list = actions.map((a) => `- ${a.name}${a.description ? ` — ${a.description}` : ""}`).join("\n");
  return NATIVE_ACTION_INSTRUCTIONS.replace("{{ACTIONS}}", list || "- (none)");
}

export const ACTION_CONFIRM_MESSAGE_TYPES = new Set([
  "ConfirmationCard",
  "TriggerConfirmation",
  "RenderCardRequest",
]);

export const ACTION_ALLOWED_MESSAGE_TYPES = [
  "TriggerConfirmation",
  "TriggerPlugin",
  "TriggerUserInputRequest",
  "TriggerExtension",
  "CompleteExtension",
  "TriggerPluginAuth",
  "HintInvocation",
];

interface AdaptiveCardAction {
  type?: string;
  title?: string;
  data?: { message?: { actionId?: string; confirmationOption?: string }; [k: string]: unknown };
  [k: string]: unknown;
}

interface AdaptiveCard {
  actions?: AdaptiveCardAction[];
  body?: Array<{ actions?: AdaptiveCardAction[]; [k: string]: unknown }>;
  [k: string]: unknown;
}

export interface MaybeTriggerMessage {
  messageType?: string;
  layout?: string;
  copilotMessageType?: string;
  actionId?: string;
  sourceRequestId?: string;
  requestId?: string;
  messageId?: string;
  isConsequential?: boolean;
  confirmationMetadata?: unknown;
  adaptiveCards?: AdaptiveCard[];
  [k: string]: unknown;
}

export interface ActionConfirmation {
  actionId: string;
  sourceRequestId: string;
  isConsequential: boolean;
  confirmationOption?: string;
  confirmationTitle?: string;
  triggerMessage: MaybeTriggerMessage;
}

function collectActions(cards: AdaptiveCard[] | undefined): AdaptiveCardAction[] {
  const out: AdaptiveCardAction[] = [];
  for (const c of cards ?? []) {
    for (const a of c.actions ?? []) out.push(a);
    for (const b of c.body ?? []) for (const a of b.actions ?? []) out.push(a);
  }
  return out;
}

const NEGATIVE = /cancel|decline|deny|reject|no\b|dismiss|not now/i;
const AFFIRMATIVE = /allow|confirm|continue|proceed|yes|approve|run|send|ok/i;

function pickAffirmative(actions: AdaptiveCardAction[]): { message: { actionId?: string; confirmationOption?: string }; title?: string } | null {
  const submits = actions.filter((a) => (a.type ?? "").includes("Submit") && a.data?.message);
  if (!submits.length) return null;
  const affirmative = submits.find((a) => AFFIRMATIVE.test(a.title ?? "") && !NEGATIVE.test(a.title ?? ""));
  const nonNegative = submits.find((a) => !NEGATIVE.test(a.title ?? ""));
  const chosen = affirmative ?? nonNegative ?? submits[0];
  return { message: chosen.data!.message ?? {}, title: chosen.title };
}

export function parseActionConfirmation(m: MaybeTriggerMessage | null | undefined): ActionConfirmation | null {
  if (!m || typeof m !== "object") return null;
  const isConfirmType =
    (m.messageType && ACTION_CONFIRM_MESSAGE_TYPES.has(m.messageType)) ||
    m.layout === "confirmation_trigger" ||
    (m.copilotMessageType === "adaptiveCard" && !!m.confirmationMetadata);

  const picked = pickAffirmative(collectActions(m.adaptiveCards));
  const actionId = m.actionId ?? picked?.message.actionId;
  if (!isConfirmType && !(actionId && m.adaptiveCards?.length)) return null;
  if (!actionId) return null;

  const sourceRequestId = m.sourceRequestId ?? m.requestId ?? m.messageId ?? "";
  return {
    actionId,
    sourceRequestId,
    isConsequential: m.isConsequential ?? true,
    confirmationOption: picked?.message.confirmationOption,
    confirmationTitle: picked?.title,
    triggerMessage: m,
  };
}

export interface ResumeInvokeOptions {
  author?: string;
  text?: string;
}

export function buildResumeInvokeAction(conf: ActionConfirmation, opts: ResumeInvokeOptions = {}): Record<string, unknown> {
  return {
    text: opts.text ?? conf.confirmationTitle ?? conf.confirmationOption ?? "confirmation response",
    messageType: "ResumeInvokeAction",
    sourceRequestId: conf.sourceRequestId,
    actionId: conf.actionId,
    author: opts.author ?? "user",
    invokeActionMessages: [conf.triggerMessage],
  };
}

export function shouldAutoConfirm(conf: ActionConfirmation, env: { autoConfirmAll?: boolean } = {}): boolean {
  if (env.autoConfirmAll) return true;
  return !conf.isConsequential;
}
