export interface AttributeRequest {
  t: string;
  v?: string;
}

export type AttributeCon = AttributeRequest[];

export interface PolicyEntry {
  ts: number;
  con: AttributeCon;
}

export interface Policy {
  [recipientId: string]: AttributeCon;
}

export interface SealPolicy {
  [recipientId: string]: PolicyEntry;
}

export type KeySort = "Decryption" | "Signing";

export interface PopupData {
  hostname: string;
  header: Record<string, string>;
  con: AttributeCon;
  sort: KeySort;
  hints?: AttributeCon;
  senderId?: string;
}

export interface Badge {
  type: string;
  value: string;
}

export interface ComposeState {
  encrypt: boolean;
  policy?: Policy;
  signId?: Policy;
}

export interface DecryptedMessageInfo {
  badges?: Badge[];
  subject?: string;
  body?: string;
  attachments?: { name: string; content: string }[];
}
