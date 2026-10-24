import { getOrCreateAgent } from "../io/agent.js";
import { CopilotSession } from "../protocol/session.js";
import { uploadImage, type ImageAnnotation } from "../io/image-input.js";
import { createLogger, trunc } from "../io/log.js";
import { awaitTurnBudget } from "../budget/turn-budget.js";
import type { CopilotStream } from "../routing/copilot.js";
import { agentAllowed, imagesAllowed, selectCredential, turnContextFor, type CredentialSelection } from "./credential.js";

const log = createLogger("model");

export interface ModelSessionOptions {
  accountId?: string | null;
  useAgent?: boolean;
  temporaryChat?: boolean;
}

export class ModelSession {
  readonly accountId: string | null;
  private useAgent: boolean;
  private temporaryChat: boolean;
  private selection: CredentialSelection | null = null;
  private copilotSession: CopilotSession | null = null;
  private cachedAgentId: string | null | undefined = undefined;

  readonly sessionId: string = crypto.randomUUID();
  private _conversationId: string = crypto.randomUUID();
  get conversationId(): string { return this._conversationId; }
  newConversation(): void {
    this._conversationId = crypto.randomUUID();
    this.reset();
  }

  constructor(options: ModelSessionOptions = {}) {
    this.accountId = options.accountId ?? null;
    this.useAgent = options.useAgent !== false;
    this.temporaryChat =
      options.temporaryChat ?? process.env.M365_SAVE_HISTORY !== "1";
  }

  get turnCount(): number {
    return this.copilotSession?.turnCount ?? 0;
  }

  private currentAgentId: string | undefined = undefined;

  private createCopilotSession(
    agentId: string | undefined,
    selection: CredentialSelection,
    model: string,
  ): CopilotSession {
    return new CopilotSession({
      agentId,
      sessionId: this.sessionId,
      conversationId: this.conversationId,
      temporaryChat: this.temporaryChat,
      accountId: this.accountId,
      model,
      turnContext: (plan) => turnContextFor(selection, plan),
    });
  }

  async run(
    text: string,
    model: string = "m365-copilot",
    signal?: AbortSignal,
    useAgent: boolean = true,
    images: string[] = [],
    waitForBudget: () => Promise<{ waitedMs: number }> = () => awaitTurnBudget({ signal }),
  ): Promise<CopilotStream> {
    this.selection = await selectCredential(this.accountId);
    let token = this.selection.token;
    const annotations: ImageAnnotation[] = [];
    for (const url of images) {
      try {
        annotations.push(await uploadImage(token, this.conversationId, url));
      } catch (err: any) {
        log.info(`Image upload failed: ${err.message}`);
        text += `

(Note: an attached image could not be passed on — ${err.message.slice(0, 120)}. Don't describe it; say it didn't arrive.)`;
      }
    }
    const wantAgent = agentAllowed(this.selection, this.useAgent && useAgent);

    if (wantAgent && this.cachedAgentId === undefined) {
      try {
        this.cachedAgentId = await getOrCreateAgent();
        if (this.cachedAgentId) log.info(`Using agent: ${this.cachedAgentId}`);
        else log.info("No agent available");
      } catch {
        this.cachedAgentId = null;
      }
    }
    const agentForTurn = wantAgent ? (this.cachedAgentId ?? undefined) : undefined;

    const generateImages = imagesAllowed(this.selection, agentForTurn);
    const turnOpts = { generateImages, images: annotations };

    if (!this.copilotSession || this.currentAgentId !== agentForTurn) {
      this.copilotSession = this.createCopilotSession(agentForTurn, this.selection, model);
      this.currentAgentId = agentForTurn;
    }

    log.info(`run: model=${model}, agent=${agentForTurn ?? "none"}, turn=${this.copilotSession.turnCount}, sid=${this.sessionId}, cid=${this.conversationId}, text=${JSON.stringify(trunc(text, 200))}`);

    try {
      return await this.copilotSession.chat(token, text, model, signal, turnOpts);
    } catch (err: any) {
      log.info("Session error, reconnecting:", err.message);
      if ((await waitForBudget()).waitedMs > 0) {
        this.selection = await selectCredential(this.accountId);
        token = this.selection.token;
      }
      this.copilotSession = this.createCopilotSession(agentForTurn, this.selection, model);
      this.currentAgentId = agentForTurn;
      return await this.copilotSession.chat(token, text, model, signal, turnOpts);
    }
  }

  async refreshAgent(): Promise<boolean> {
    if (!this.useAgent) return false;
    try {
      const fresh = await getOrCreateAgent({ forceRefresh: true });
      if (fresh !== this.cachedAgentId) {
        log.info(`Agent refreshed: ${this.cachedAgentId ?? "none"} -> ${fresh ?? "none"}`);
        this.cachedAgentId = fresh;
        this.copilotSession = null;
        this.currentAgentId = undefined;
        return true;
      }
    } catch (err: any) {
      log.info(`Agent refresh failed: ${err.message}`);
    }
    return false;
  }

  reset() {
    this.copilotSession = null;
    this.cachedAgentId = undefined;
    this.currentAgentId = undefined;
  }
}
