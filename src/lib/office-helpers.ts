// Promisified wrappers around the callback-based Office.js mailbox APIs.
// All helpers reject if Office.js returns a non-succeeded AsyncResult.

function p<T>(fn: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    fn((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value);
      } else {
        reject(res.error);
      }
    });
  });
}

export function getItem(): Office.MessageCompose | Office.MessageRead {
  return Office.context.mailbox.item as Office.MessageCompose | Office.MessageRead;
}

export function isComposeMode(): boolean {
  const item = getItem() as { subject?: unknown };
  // In compose mode item.subject is a Subject object (with getAsync/
  // setAsync). In read mode item.subject is the message subject string.
  // We previously narrowed by `typeof setAsync === 'function'`, but
  // Outlook for Mac's Hx-rewrite layer wraps Office.js methods in
  // proxies that fail that strict check. Compare against `string`
  // instead — anything non-string is the compose-mode object.
  return typeof item.subject !== "string";
}

// --- Compose getters ---

export function getSubject(): Promise<string> {
  const item = getItem() as Office.MessageCompose;
  return p<string>((cb) => item.subject.getAsync(cb));
}

export function setSubject(subject: string): Promise<void> {
  const item = getItem() as Office.MessageCompose;
  return p<void>((cb) => item.subject.setAsync(subject, cb));
}

export function getRecipients(field: "to" | "cc" | "bcc"): Promise<Office.EmailAddressDetails[]> {
  const item = getItem() as Office.MessageCompose;
  const recipientsField = item[field];
  return p<Office.EmailAddressDetails[]>((cb) => recipientsField.getAsync(cb));
}

export interface BodyResult {
  body: string;
  format: Office.CoercionType;
}

export function getBody(format: Office.CoercionType = Office.CoercionType.Html): Promise<string> {
  const item = getItem() as Office.MessageCompose;
  return p<string>((cb) => item.body.getAsync(format, cb));
}

export function setBody(html: string): Promise<void> {
  const item = getItem() as Office.MessageCompose;
  return p<void>((cb) => item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, cb));
}

export function getAttachmentsCompose(): Promise<Office.AttachmentDetailsCompose[]> {
  const item = getItem() as Office.MessageCompose;
  return p<Office.AttachmentDetailsCompose[]>((cb) => item.getAttachmentsAsync(cb));
}

export function getAttachmentContentCompose(
  attachmentId: string
): Promise<Office.AttachmentContent> {
  const item = getItem() as Office.MessageCompose;
  return p<Office.AttachmentContent>((cb) => item.getAttachmentContentAsync(attachmentId, cb));
}

export function removeAttachment(attachmentId: string): Promise<void> {
  const item = getItem() as Office.MessageCompose;
  return p<void>((cb) => item.removeAttachmentAsync(attachmentId, cb));
}

// Adds a base64 inline file attachment. Returns the attachment id assigned by Outlook.
export function addBase64Attachment(filename: string, base64: string): Promise<string> {
  const item = getItem() as Office.MessageCompose;
  return p<string>((cb) => item.addFileAttachmentFromBase64Async(base64, filename, cb as never));
}

// Commits the current draft (subject, body, attachments) to the server. We need
// this after writing the encrypted body + attachment because Send otherwise races
// the server-side upload of those changes; new Outlook on Windows shows a Smart
// Alerts-style "PostGuard timed out" dialog when that race occurs.
export function saveItem(): Promise<string> {
  const item = getItem() as Office.MessageCompose;
  return p<string>((cb) => item.saveAsync(cb));
}

// Internet header storage. Used to share state (e.g. the encrypt toggle)
// between the taskpane and the OnMessageSend launch event handler, which
// runs in a separate runtime. customProperties does not propagate cross-
// runtime in new Outlook (OWA-based), but internet headers are persisted
// onto the message itself so they're guaranteed visible on send.
//
// Custom internet header names must start with "x-" per Office.js.
export function setItemHeaders(headers: Record<string, string>): Promise<void> {
  const item = getItem() as Office.MessageCompose;
  return p<void>((cb) => item.internetHeaders.setAsync(headers, cb));
}

export function removeItemHeaders(names: string[]): Promise<void> {
  const item = getItem() as Office.MessageCompose;
  return p<void>((cb) => item.internetHeaders.removeAsync(names, cb));
}

export function getItemHeaders(names: string[]): Promise<Record<string, string>> {
  const item = getItem() as Office.MessageCompose;
  return p<Record<string, string>>((cb) => item.internetHeaders.getAsync(names, cb));
}

// --- Read mode getters ---

export function getReadAttachments(): Office.AttachmentDetails[] {
  const item = getItem() as Office.MessageRead;
  return item.attachments ?? [];
}

export function getReadBody(
  format: Office.CoercionType = Office.CoercionType.Html
): Promise<string> {
  const item = getItem() as Office.MessageRead;
  return p<string>((cb) => item.body.getAsync(format, cb));
}

export function getReadSubject(): string {
  const item = getItem() as Office.MessageRead;
  return item.subject as string;
}

export function getReadFrom(): Office.EmailAddressDetails | undefined {
  const item = getItem() as Office.MessageRead;
  return item.from;
}

export function getReadToRecipients(): Office.EmailAddressDetails[] {
  const item = getItem() as Office.MessageRead;
  return item.to ?? [];
}

export function getReadCcRecipients(): Office.EmailAddressDetails[] {
  const item = getItem() as Office.MessageRead;
  return item.cc ?? [];
}

export function getReadInternetHeaders(names: string[]): Promise<Record<string, string>> {
  const item = getItem() as Office.MessageRead;
  return p<string>((cb) => item.getAllInternetHeadersAsync?.(cb))
    .then((raw) => parseInternetHeaders(raw ?? "", names))
    .catch(() => ({}));
}

function parseInternetHeaders(raw: string, names: string[]): Record<string, string> {
  const wanted = new Set(names.map((n) => n.toLowerCase()));
  const out: Record<string, string> = {};
  if (!raw) return out;
  for (const line of raw.split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx < 0) continue;
    const name = line.slice(0, idx).trim().toLowerCase();
    if (!wanted.has(name)) continue;
    out[name] = line.slice(idx + 1).trim();
  }
  return out;
}

// Reads attachment bytes (compose mode). Returns ArrayBuffer.
export async function readComposeAttachmentBytes(attachmentId: string): Promise<ArrayBuffer> {
  const content = await getAttachmentContentCompose(attachmentId);
  if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    return base64ToArrayBuffer(content.content);
  }
  // Fallback: treat as utf-8 string.
  return new TextEncoder().encode(content.content).buffer as ArrayBuffer;
}

function base64ToArrayBuffer(b64: string): ArrayBuffer {
  const bin = atob(b64);
  const buf = new ArrayBuffer(bin.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < bin.length; i++) view[i] = bin.charCodeAt(i);
  return buf;
}

// Reads attachment bytes (read mode) by REST or makeEwsRequest fallback.
export function getReadAttachmentContent(attachmentId: string): Promise<Office.AttachmentContent> {
  const item = getItem() as Office.MessageRead;
  // getAttachmentContentAsync also exists in read mode.
  return p<Office.AttachmentContent>((cb) =>
    (
      item as unknown as {
        getAttachmentContentAsync: (
          id: string,
          cb: (r: Office.AsyncResult<Office.AttachmentContent>) => void
        ) => void;
      }
    ).getAttachmentContentAsync(attachmentId, cb)
  );
}

export async function readReadAttachmentBytes(attachmentId: string): Promise<ArrayBuffer> {
  const content = await getReadAttachmentContent(attachmentId);
  if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    return base64ToArrayBuffer(content.content);
  }
  return new TextEncoder().encode(content.content).buffer as ArrayBuffer;
}

export function getSenderEmail(): string {
  const profile = Office.context.mailbox.userProfile;
  return profile.emailAddress.toLowerCase();
}

export function getSenderDisplayName(): string {
  return Office.context.mailbox.userProfile.displayName || "";
}

export interface NotificationOptions {
  type?: "informational" | "error";
  persistent?: boolean;
}

export function showNotification(
  key: string,
  message: string,
  opts: NotificationOptions = {}
): Promise<void> {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return Promise.resolve();
  const details: Office.NotificationMessageDetails = {
    type:
      opts.type === "error"
        ? Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage
        : Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message.slice(0, 150),
    icon: "Icon.16x16",
    persistent: opts.persistent ?? false,
  };
  return p<void>((cb) => item.notificationMessages.replaceAsync(key, details, cb));
}
