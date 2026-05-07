// Read-mode taskpane view: detect PostGuard envelope, decrypt with Yivi
// inline, and render the plaintext + sender badges in the taskpane.
//
// Outlook does not let an add-in modify the displayed message, so the
// decrypted content is shown only inside this taskpane and a small
// notification banner is added to the message.

import { PostGuard } from "@e4a/pg-js";
import {
  getReadAttachments,
  readReadAttachmentBytes,
  getReadBody,
  getReadFrom,
  getReadToRecipients,
  getReadCcRecipients,
  getSenderEmail,
  showNotification,
} from "../lib/office-helpers";
import { fromBase64, bytesToUtf8 } from "../lib/encoding";
import {
  POSTGUARD_ENCRYPTED_FILENAME,
  extractArmoredCiphertext,
  looksLikePostGuard,
  parseDecryptedMime,
  ParsedAttachment,
  readMimeHeader,
} from "../lib/mime";
import { Badge, FriendlySender } from "../lib/types";
import { PKG_URL, CRYPTIFY_URL, clientHeaders } from "../lib/pkg-client";
import { t } from "../lib/i18n";
import { showView, setStatus, showError } from "./taskpane";

const ADDIN_VERSION = "0.1.0";

interface ReadState {
  ciphertext: Uint8Array | null;
  recipientEmail: string;
  busy: boolean;
}

const state: ReadState = {
  ciphertext: null,
  recipientEmail: "",
  busy: false,
};

export async function mountReadView(): Promise<void> {
  state.recipientEmail = pickRecipientEmail();

  const ciphertext = await tryFindCiphertext();
  if (ciphertext) {
    state.ciphertext = ciphertext;
    showEncryptedView();
    await showNotification("postguard-encrypted-banner", t("displayScriptDecryptBar"), {
      type: "informational",
      persistent: true,
    });
    return;
  }

  // Was this message originally encrypted (decrypted earlier)?
  const wasEncrypted = await checkWasEncrypted();
  if (wasEncrypted) {
    const text = byId<HTMLElement>("pg-was-encrypted-text");
    text.textContent = t("displayScriptWasEncryptedBar");
    showView("read_was_encrypted");
    return;
  }

  showView("read_noop");
}

function showEncryptedView(): void {
  const text = byId<HTMLElement>("pg-read-encrypted-text");
  const btn = byId<HTMLButtonElement>("pg-btn-decrypt");
  text.textContent = t("displayScriptDecryptBar");
  btn.textContent = t("decryptButton");

  // Replace listeners by cloning.
  const fresh = btn.cloneNode(true) as HTMLButtonElement;
  btn.replaceWith(fresh);
  fresh.addEventListener("click", () => {
    if (state.busy) return;
    void runDecryption();
  });

  showView("read_encrypted");
}

async function tryFindCiphertext(): Promise<Uint8Array | null> {
  // Path 1: postguard.encrypted attachment.
  const attachments = getReadAttachments();
  const enc = attachments.find((a) => a.name?.toLowerCase() === POSTGUARD_ENCRYPTED_FILENAME);
  if (enc) {
    try {
      const buf = await readReadAttachmentBytes(enc.id);
      return new Uint8Array(buf);
    } catch (_e) {
      // Fall through to body-armor fallback.
    }
  }

  // Path 2: ASCII-armored block in the body.
  try {
    const html = await getReadBody(Office.CoercionType.Html);
    const armored = extractArmoredCiphertext(html);
    if (armored) {
      return fromBase64(armored);
    }
    if (looksLikePostGuard(html)) {
      // Armor markers were present but content not extractable — still
      // treat as encrypted so the user gets an error instead of silence.
      return new Uint8Array();
    }
  } catch (_e) {
    // Ignore.
  }

  return null;
}

async function checkWasEncrypted(): Promise<boolean> {
  // Read mode does not give us trailers without makeEwsRequest/Graph.
  // For the cheap check we look at the visible body for our marker.
  try {
    const html = await getReadBody(Office.CoercionType.Html);
    return /postguard\.encrypted|x-postguard/i.test(html);
  } catch (_e) {
    return false;
  }
}

function pickRecipientEmail(): string {
  // Prefer the active mailbox account email. Falls back to the first
  // To/Cc address — relevant when the message was sent to a shared
  // mailbox or alias.
  const own = getSenderEmail();
  if (own) return own;
  const to = getReadToRecipients();
  if (to.length > 0) return to[0].emailAddress.toLowerCase();
  const cc = getReadCcRecipients();
  if (cc.length > 0) return cc[0].emailAddress.toLowerCase();
  return "";
}

async function runDecryption(): Promise<void> {
  if (!state.ciphertext || state.ciphertext.length === 0) {
    showError(t("decryptionError"));
    return;
  }
  if (!state.recipientEmail) {
    showError(t("recipientUnknown"));
    return;
  }

  state.busy = true;
  setStatus(t("decryptingButton"));
  try {
    showView("yivi");
    const yiviTitle = byId<HTMLElement>("pg-yivi-title");
    const yiviSubtitle = byId<HTMLElement>("pg-yivi-subtitle");
    yiviTitle.textContent = `${t("displayMessageTitle")} ${getReadFrom()?.emailAddress ?? ""}`;
    yiviSubtitle.textContent = t("displayMessageQrPrefix");
    document.getElementById("yivi-web-form")!.innerHTML = "";

    const pg = new PostGuard({
      pkgUrl: PKG_URL,
      cryptifyUrl: CRYPTIFY_URL,
      headers: clientHeaders(ADDIN_VERSION),
    } as never);

    const opened = (
      pg as never as {
        open: (input: { data: Uint8Array }) => OpenedMessage;
      }
    ).open({ data: state.ciphertext });

    const result = await opened.decrypt({
      element: "#yivi-web-form",
      recipient: state.recipientEmail,
    });

    renderDecrypted(result.plaintext, result.sender);
    setStatus("");
  } catch (err) {
    const message = err instanceof Error ? err.message : t("decryptionError");
    if (/KEM/i.test(message)) {
      showError(t("decryptionFailed"));
    } else {
      showError(message);
    }
  } finally {
    state.busy = false;
  }
}

interface OpenedMessage {
  decrypt(opts: { element: string; recipient: string }): Promise<{
    plaintext: Uint8Array;
    sender: FriendlySender | null;
  }>;
}

function renderDecrypted(plaintext: Uint8Array, sender: FriendlySender | null): void {
  const mime = bytesToUtf8(plaintext);
  const subject = readMimeHeader(mime, "Subject") ?? "";
  const from = readMimeHeader(mime, "From") ?? "";
  const date = readMimeHeader(mime, "Date") ?? "";

  const subjectEl = byId<HTMLElement>("pg-decrypted-subject");
  subjectEl.textContent = subject;

  const metaEl = byId<HTMLElement>("pg-decrypted-meta");
  metaEl.textContent = [from && `From: ${from}`, date && `Date: ${date}`]
    .filter(Boolean)
    .join("  •  ");

  const badgesEl = byId<HTMLElement>("pg-decrypted-badges");
  badgesEl.innerHTML = "";
  const badges = badgesFromSender(sender);
  if (badges.length > 0) {
    const label = document.createElement("span");
    label.textContent = `${t("notificationHeaderBadgesLabel")}: `;
    label.className = "pg-meta";
    badgesEl.appendChild(label);
    for (const b of badges) {
      const span = document.createElement("span");
      span.className = "pg-badge";
      span.textContent = b.value;
      badgesEl.appendChild(span);
    }
  }

  const parsed = parseDecryptedMime(mime);
  const bodyText = parsed.htmlBody ?? parsed.plainBody ?? "";
  const isHtml = parsed.htmlBody != null;

  const iframe = byId<HTMLIFrameElement>("pg-decrypted-body");
  iframe.srcdoc = wrapHtml(bodyText, isHtml);

  renderAttachments(parsed.attachments);

  showView("decrypted");
}

let attachmentObjectUrls: string[] = [];

function renderAttachments(attachments: ParsedAttachment[]): void {
  // Revoke any blobs from a previous decryption to free memory.
  for (const url of attachmentObjectUrls) URL.revokeObjectURL(url);
  attachmentObjectUrls = [];

  const host = byId<HTMLElement>("pg-decrypted-attachments");
  host.innerHTML = "";
  if (attachments.length === 0) {
    host.hidden = true;
    return;
  }
  host.hidden = false;

  const heading = document.createElement("h4");
  heading.className = "pg-meta pg-attachments-heading";
  heading.textContent = `${t("decryptedAttachmentsHeading", "Attachments")} (${attachments.length})`;
  host.appendChild(heading);

  const list = document.createElement("ul");
  list.className = "pg-attachment-list";
  for (const att of attachments) {
    const li = document.createElement("li");
    const blob = new Blob([att.data as BlobPart], { type: att.type || "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    attachmentObjectUrls.push(url);
    const a = document.createElement("a");
    a.href = url;
    a.download = att.name;
    a.textContent = att.name;
    a.className = "pg-attachment-link";
    const size = document.createElement("span");
    size.className = "pg-meta";
    size.textContent = `  (${formatSize(att.data.byteLength)})`;
    li.appendChild(a);
    li.appendChild(size);
    list.appendChild(li);
  }
  host.appendChild(list);
}

function formatSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function badgesFromSender(sender: FriendlySender | null): Badge[] {
  if (!sender) return [];
  const out: Badge[] = [];
  if (sender.email) out.push({ value: sender.email });
  for (const a of sender.attributes ?? []) {
    if (a.value) out.push({ value: a.value });
  }
  return out;
}

function wrapHtml(body: string, isHtml: boolean): string {
  if (isHtml) {
    if (/<html[\s>]|<body[\s>]/i.test(body)) return body;
    return `<!doctype html><html><body>${body}</body></html>`;
  }
  return `<!doctype html><html><body><pre style="white-space:pre-wrap;font-family:Segoe UI,Helvetica,Arial,sans-serif">${escape(
    body
  )}</pre></body></html>`;
}

function escape(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
}
