/* global Office */

import type { ISealOptions } from "@e4a/pg-wasm";
import {
  PG_ATTACHMENT_NAME,
  EMAIL_ATTRIBUTE_TYPE,
  POSTGUARD_SUBJECT,
  toEmail,
  retrievePublicKey,
  retrieveVerificationKey,
  getUSK,
  getSigningKeys,
  checkJwtCache,
  PG_CLIENT_HEADER,
  PKG_URL,
  secondsTill4AM,
  buildEncryptedBody,
  extractArmoredPayload,
  parseMimeContent,
} from "../utils";
import type { AttributeCon, Policy, SealPolicy, ComposeState } from "../types";

const PG_HEADER_NAME = "x-postguard";
const PG_HEADER_VALUE = "0.1.0";

Office.onReady(() => {
  console.log("[PostGuard LaunchEvent] Office.onReady fired");
});

// ─── Compose helpers (for OnMessageSend) ───────────────────────────

function getComposeState(): ComposeState {
  const saved = sessionStorage.getItem("pg-compose-state");
  if (!saved) return { encrypt: false };
  try {
    return JSON.parse(saved);
  } catch {
    return { encrypt: false };
  }
}

async function getBody(item: Office.MessageCompose): Promise<{ body: string; isHtml: boolean }> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve({ body: result.value, isHtml: true });
      } else {
        item.body.getAsync(Office.CoercionType.Text, (textResult) => {
          if (textResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve({ body: textResult.value, isHtml: false });
          } else {
            reject(new Error("Failed to get message body"));
          }
        });
      }
    });
  });
}

async function getSubject(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve) => {
    item.subject.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : "");
    });
  });
}

async function getRecipients(item: Office.MessageCompose): Promise<string[]> {
  const toPromise = new Promise<Office.EmailAddressDetails[]>((resolve) => {
    item.to.getAsync((r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : [])
    );
  });
  const ccPromise = new Promise<Office.EmailAddressDetails[]>((resolve) => {
    item.cc.getAsync((r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : [])
    );
  });
  const [toList, ccList] = await Promise.all([toPromise, ccPromise]);
  return [...toList, ...ccList].map((r) => toEmail(r.emailAddress));
}

async function getSender(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve) => {
    item.from.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(toEmail(result.value.emailAddress));
      } else {
        resolve(Office.context.mailbox.userProfile.emailAddress.toLowerCase());
      }
    });
  });
}

async function setBody(item: Office.MessageCompose, content: string, isHtml: boolean): Promise<void> {
  return new Promise((resolve, reject) => {
    const coercionType = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
    item.body.setAsync(content, { coercionType }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error("Failed to set body"));
    });
  });
}

async function setSubject(item: Office.MessageCompose, subject: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.subject.setAsync(subject, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error("Failed to set subject"));
    });
  });
}

async function addAttachment(
  item: Office.MessageCompose,
  base64: string,
  name: string,
  type: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    item.addFileAttachmentFromBase64Async(base64, name, { asyncContext: type }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error("Failed to add attachment"));
    });
  });
}

async function setInternetHeader(
  item: Office.MessageCompose,
  name: string,
  value: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    item.internetHeaders.setAsync({ [name]: value }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error("Failed to set internet header"));
    });
  });
}

// ─── OnMessageSend handler ─────────────────────────────────────────

async function onMessageSendHandler(event: Office.AddinCommands.Event): Promise<void> {
  console.log("[PostGuard] onMessageSendHandler triggered");

  const state = getComposeState();
  console.log("[PostGuard] Compose state:", JSON.stringify(state));

  if (!state.encrypt) {
    console.log("[PostGuard] Encryption not enabled, allowing send");
    event.completed({ allowEvent: true });
    return;
  }

  try {
    const item = Office.context.mailbox.item as unknown as Office.MessageCompose;
    console.log("[PostGuard] Got mailbox item");

    console.log("[PostGuard] Fetching message details...");
    const [originalSubject, { body, isHtml }, recipients, from] = await Promise.all([
      getSubject(item),
      getBody(item),
      getRecipients(item),
      getSender(item),
    ]);
    console.log("[PostGuard] Subject:", originalSubject);
    console.log("[PostGuard] Body length:", body.length, "isHtml:", isHtml);
    console.log("[PostGuard] Recipients:", recipients);
    console.log("[PostGuard] From:", from);

    if (recipients.length === 0) {
      console.log("[PostGuard] No recipients, blocking send");
      item.notificationMessages.replaceAsync("pgError", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "PostGuard: No recipients found.",
      });
      event.completed({ allowEvent: false });
      return;
    }

    console.log("[PostGuard] Loading WASM module and public key...");
    const [pk, mod] = await Promise.all([retrievePublicKey(), import("@e4a/pg-wasm")]);
    // Initialize the WASM module (default export is the init function)
    await mod.default();
    console.log("[PostGuard] WASM and public key loaded");

    const timestamp = Math.round(Date.now() / 1000);
    const customPolicies = state.policy;

    const policy: SealPolicy = {};
    for (const recipientEmail of recipients) {
      if (customPolicies && customPolicies[recipientEmail] && customPolicies[recipientEmail].length > 0) {
        policy[recipientEmail] = { ts: timestamp, con: customPolicies[recipientEmail] };
      } else {
        policy[recipientEmail] = {
          ts: timestamp,
          con: [{ t: EMAIL_ATTRIBUTE_TYPE, v: recipientEmail }],
        };
      }
      policy[recipientEmail].con = policy[recipientEmail].con.map(({ t, v }) => {
        if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: (v || "").toLowerCase() };
        return { t, v };
      });
    }
    console.log("[PostGuard] Policy built:", JSON.stringify(policy));

    const pubSignId: AttributeCon = [{ t: EMAIL_ATTRIBUTE_TYPE, v: from }];

    console.log("[PostGuard] Getting signing JWT...");
    const jwt = await checkJwtCache(pubSignId).catch((e) => {
      console.log("[PostGuard] No cached JWT:", e.message);
      throw new Error("No cached signing JWT. Please configure your signing identity in the compose pane before sending.");
    });
    console.log("[PostGuard] Got signing JWT");

    console.log("[PostGuard] Getting signing keys...");
    const { pubSignKey, privSignKey } = await getSigningKeys(jwt, { pubSignId });
    console.log("[PostGuard] Got signing keys, pubSignKey:", !!pubSignKey, "privSignKey:", !!privSignKey);

    const sealOptions: ISealOptions = {
      policy,
      pubSignKey,
      ...(privSignKey && { privSignKey }),
    };

    const date = new Date();
    const contentType = isHtml ? "text/html; charset=utf-8" : "text/plain; charset=utf-8";

    let innerMime = "";
    innerMime += `Date: ${date.toUTCString()}\r\n`;
    innerMime += "MIME-Version: 1.0\r\n";
    innerMime += `To: ${recipients.join(", ")}\r\n`;
    innerMime += `From: ${from}\r\n`;
    innerMime += `Subject: ${originalSubject}\r\n`;
    innerMime += `Content-Type: ${contentType}\r\n`;
    innerMime += `X-PostGuard: 0.1\r\n`;
    innerMime += "\r\n";
    innerMime += body;
    console.log("[PostGuard] Inner MIME built, length:", innerMime.length);

    const encoder = new TextEncoder();
    const readable = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(encoder.encode(innerMime));
        controller.close();
      },
    });

    let encrypted = new Uint8Array(0);
    const writable = new WritableStream<Uint8Array>({
      write(chunk: Uint8Array) {
        const combined = new Uint8Array(encrypted.length + chunk.length);
        combined.set(encrypted);
        combined.set(chunk, encrypted.length);
        encrypted = combined;
      },
    });

    console.log("[PostGuard] Sealing...");
    const tStart = performance.now();
    await mod.sealStream(pk, sealOptions, readable, writable);
    console.log(`[PostGuard] Encryption took ${performance.now() - tStart} ms, size: ${encrypted.length} bytes`);

    let binary = "";
    for (let i = 0; i < encrypted.length; i++) {
      binary += String.fromCharCode(encrypted[i]);
    }
    const base64Encrypted = btoa(binary);
    console.log("[PostGuard] Base64 encoded, length:", base64Encrypted.length);

    const encryptedBody = buildEncryptedBody(from, base64Encrypted);

    console.log("[PostGuard] Setting subject...");
    await setSubject(item, POSTGUARD_SUBJECT);
    console.log("[PostGuard] Setting body...");
    await setBody(item, encryptedBody, true);

    console.log("[PostGuard] Adding encrypted attachment (best-effort)...");
    try {
      await addAttachment(item, base64Encrypted, PG_ATTACHMENT_NAME, "application/postguard");
    } catch (e) {
      console.warn("[PostGuard] Attachment failed, body fallback in place:", e);
    }

    console.log("[PostGuard] Setting internet header...");
    await setInternetHeader(item, PG_HEADER_NAME, PG_HEADER_VALUE);

    console.log("[PostGuard] All done, allowing send");
    event.completed({ allowEvent: true });
  } catch (e: unknown) {
    console.error("[PostGuard] Encryption error:", e);
    console.error("[PostGuard] Error stack:", e instanceof Error ? e.stack : "no stack");
    const item = Office.context.mailbox.item;
    if (item) {
      item.notificationMessages.replaceAsync("pgError", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: `PostGuard encryption failed: ${e instanceof Error ? e.message : "Unknown error"}`,
      });
    }
    event.completed({ allowEvent: false });
  }
}

// ─── OnMessageRead handler ─────────────────────────────────────────

async function getReadBody(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error("Failed to get body"));
    });
  });
}

async function decryptBytes(
  bytes: Uint8Array,
  mod: typeof import("@e4a/pg-wasm"),
  vk: string,
  event: Office.AddinCommands.Event
): Promise<void> {
  const readable = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(bytes);
      controller.close();
    },
  });

  const unsealer = await mod.StreamUnsealer.new(readable, vk);
  const userEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  const recipients = unsealer.inspect_header();
  const me = recipients.get(userEmail);

  if (!me) {
    console.log("[PostGuard] User not in recipients:", userEmail);
    event.completed({ allowEvent: false });
    return;
  }

  const keyRequest = { ...me };
  let hints: AttributeCon = me.con;

  hints = hints.map(({ t, v }) => {
    if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: userEmail };
    return { t, v };
  });

  keyRequest.con = keyRequest.con.map(({ t, v }: { t: string; v: string }) => {
    if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: userEmail };
    if (v === "" || v?.includes("*")) return { t };
    return { t, v };
  });

  let jwt: string;
  try {
    jwt = await checkJwtCache(hints);
    console.log("[PostGuard] Got cached JWT");
  } catch {
    console.log("[PostGuard] No cached JWT, cannot auto-decrypt");
    event.completed({ allowEvent: false });
    return;
  }

  const usk = await getUSK(jwt, keyRequest.ts);

  let decryptedData = "";
  const decoder = new TextDecoder();
  const writableDecrypt = new WritableStream({
    write(chunk: Uint8Array) {
      decryptedData += decoder.decode(chunk, { stream: true });
    },
    close() {
      decryptedData += decoder.decode();
    },
  });

  await unsealer.unseal(userEmail, usk, writableDecrypt);
  console.log("[PostGuard] Decryption complete, length:", decryptedData.length);

  const { subject, body, isHtml } = parseMimeContent(decryptedData);

  // Prepend the original subject since event.completed() cannot set the subject field.
  let content: string;
  if (isHtml) {
    content = `<h2 style="margin:0 0 12px">${subject}</h2>${body}`;
  } else {
    content = `${subject}\n\n${body}`;
  }

  // Use the OnMessageRead event API to replace the displayed message content.
  // This is a display-only replacement; the launch event re-fires on each open.
  (event as any).completed({
    allowEvent: true,
    emailBody: {
      coercionType: isHtml ? Office.CoercionType.Html : Office.CoercionType.Text,
      content,
    },
  });
}

async function onMessageReadHandler(event: Office.AddinCommands.Event): Promise<void> {
  console.log("[PostGuard] onMessageReadHandler triggered");
  try {
    const item = Office.context.mailbox.item;
    if (!item) {
      console.log("[PostGuard] No mailbox item");
      event.completed({ allowEvent: false });
      return;
    }

    const [, vk, mod] = await Promise.all([
      retrievePublicKey(),
      retrieveVerificationKey(),
      import("@e4a/pg-wasm"),
    ]);
    await mod.default();
    console.log("[PostGuard] WASM and keys loaded");

    // Try attachment first (item.attachments is a sync property in read mode)
    const attachmentId = findEncryptedAttachment(item);
    if (attachmentId) {
      console.log("[PostGuard] Found encrypted attachment:", attachmentId);
      const base64Content = await getAttachmentContent(item, attachmentId);
      const binaryString = atob(base64Content);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      console.log("[PostGuard] Attachment decoded, size:", bytes.length);
      await decryptBytes(bytes, mod, vk, event);
      return;
    }

    // Fallback: extract armored payload from body
    console.log("[PostGuard] No attachment found, checking body for armor...");
    const bodyHtml = await getReadBody(item);
    const armoredBase64 = extractArmoredPayload(bodyHtml);
    if (armoredBase64) {
      console.log("[PostGuard] Found armored payload in body, length:", armoredBase64.length);
      const binaryString = atob(armoredBase64);
      const bytes = new Uint8Array(binaryString.length);
      for (let i = 0; i < binaryString.length; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }
      await decryptBytes(bytes, mod, vk, event);
      return;
    }

    console.log("[PostGuard] No encrypted content found");
    event.completed({ allowEvent: true });
  } catch (e: unknown) {
    console.error("[PostGuard] OnMessageRead error:", e);
    event.completed({ allowEvent: false });
  }
}

// ─── Read helpers ──────────────────────────────────────────────────

function findEncryptedAttachment(item: Office.MessageRead): string | null {
  try {
    const attachments: Office.AttachmentDetails[] = (item as any).attachments ?? [];
    const pgAttachment = attachments.find((att) => att.name === PG_ATTACHMENT_NAME);
    return pgAttachment?.id || null;
  } catch {
    return null;
  }
}

function getAttachmentContent(item: Office.MessageRead, attachmentId: string): Promise<string> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error("Failed to get attachment content"));
        return;
      }
      resolve(result.value.content);
    });
  });
}

// ─── Register handlers ─────────────────────────────────────────────

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
Office.actions.associate("onMessageReadHandler", onMessageReadHandler);
