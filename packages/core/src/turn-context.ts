import type { AccountId } from "./account-registry.js";
import type { PersonalRecord } from "./adapters/personal/personal-session.js";

export interface SessionStore {
  read(accountId: AccountId | null): PersonalRecord | null;
  isFresh(record: PersonalRecord, now: number): boolean;
}

export interface CredentialSource {
  token(accountId: AccountId | null): Promise<string>;
  record(accountId: AccountId | null, store: SessionStore): PersonalRecord | null;
}

export interface TurnContext {
  token: string;
  socketUrl: string;
  headers: Record<string, string>;
  payloadDefaults: Record<string, unknown>;
  capabilities: string[];
}
