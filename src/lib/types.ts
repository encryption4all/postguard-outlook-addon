// Shared types for the PostGuard Outlook add-in.

export interface AttributeRequest {
  t: string;
  // Required disclosure value (Manage Access / recipient policies). Omit
  // together with `optional: true` for sign-side attributes that the user
  // may disclose at session time without a pre-specified value.
  v?: string;
  optional?: boolean;
}

export type AttributeCon = AttributeRequest[];

// Maps recipient/sender email -> required attributes.
export type Policy = Record<string, AttributeCon>;

export interface Badge {
  value: string;
}

export interface SerializedRecipient {
  type: "email" | "emailDomain";
  email: string;
  policy?: AttributeRequest[];
}

export interface FriendlySender {
  email: string | null;
  attributes: { type: string; value?: string }[];
}

// Attachment data passed into the SDK's buildMime helper.
export interface MimeAttachment {
  name: string;
  type: string;
  data: ArrayBuffer;
}

export interface ComposeSnapshot {
  from: string;
  to: string[];
  cc: string[];
  bcc: string[];
  subject: string;
  htmlBody?: string;
  plainTextBody?: string;
  attachments: MimeAttachment[];
  inReplyTo?: string;
  references?: string;
  date: Date;
}

export type ItemMode = "compose" | "read";
