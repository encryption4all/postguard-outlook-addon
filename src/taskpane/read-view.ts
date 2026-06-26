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
import { fromBase64, toBase64, bytesToUtf8 } from "../lib/encoding";
import {
  POSTGUARD_ENCRYPTED_FILENAME,
  extractArmoredCiphertext,
  looksLikePostGuard,
  parseDecryptedMime,
  ParsedAttachment,
  readMimeHeader,
} from "../lib/mime";
import { Badge, FriendlySender } from "../lib/types";
import {
  PKG_URL,
  CRYPTIFY_URL,
  ADDIN_VERSION,
  ADDIN_PUBLIC_URL,
  clientHeaders,
} from "../lib/pkg-client";
import { byId } from "../lib/dom";
import { wrapHtml } from "../lib/render-body";
import { chunkPayload } from "../lib/dialog-chunk";
import { getAllowOptimisticDialog } from "../lib/settings";
import { t } from "../lib/i18n";
import { stringifyError } from "../lib/stringify-error";
import type { DecryptedMessagePayload } from "../read-dialog/read-dialog";
import { showView, setStatus, showError } from "./taskpane";

// Parsed, ready-to-render form of a decrypted message. Held in memory only
// (on `state.decrypted`) so the dialog can be re-opened without decrypting
// again; never written anywhere persistent.
interface DecryptedContent {
  subject: string;
  from: string;
  date: string;
  badges: string[];
  body: string;
  isHtml: boolean;
  attachments: ParsedAttachment[];
}

interface ReadState {
  ciphertext: Uint8Array | null;
  recipientEmail: string;
  busy: boolean;
  decrypted: DecryptedContent | null;
}

const state: ReadState = {
  ciphertext: null,
  recipientEmail: "",
  busy: false,
  decrypted: null,
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
    const detail = stringifyError(err);
    const message = detail || t("decryptionError");
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
  const parsed = parseDecryptedMime(mime);
  const content: DecryptedContent = {
    subject: readMimeHeader(mime, "Subject") ?? "",
    from: readMimeHeader(mime, "From") ?? "",
    date: readMimeHeader(mime, "Date") ?? "",
    badges: badgesFromSender(sender).map((b) => b.value),
    body: parsed.htmlBody ?? parsed.plainBody ?? "",
    isHtml: parsed.htmlBody != null,
    attachments: parsed.attachments,
  };
  // Keep the parsed plaintext in memory so the dialog can be re-opened
  // without a second Yivi/decrypt round-trip. Nothing is persisted.
  state.decrypted = content;
  void presentDecrypted();
}

// Show the decrypted message in a roomy popup dialog (issue #72). The
// cramped taskpane is a poor fit for reading a full email, so the
// already-decrypted content is marshalled into a dedicated dialog window
// in memory. If the dialog cannot be opened (popup blocked, host
// limitation), fall back to rendering inline in the taskpane so the user
// can still read their message.
async function presentDecrypted(): Promise<void> {
  const content = state.decrypted;
  if (!content) return;
  try {
    await openDecryptedDialog(content);
    showDialogOpenedView();
  } catch (err) {
    console.log(`[pg-read] dialog open failed, falling back to taskpane: ${stringifyError(err)}`);
    renderInTaskpane(content);
  }
}

// displayDialogAsync only accepts a screen percentage. Convert a target
// pixel size, clamped to Office's [1, 99] range. Mirrors the launchevent
// helper so the dialog scales sensibly on both laptops and ultrawides.
function pctOfScreen(targetPx: number, screenPx: number): number {
  const pct = Math.ceil((targetPx / screenPx) * 100);
  return Math.min(99, Math.max(1, pct));
}

function buildDialogPayload(content: DecryptedContent): DecryptedMessagePayload {
  return {
    type: "decrypted-message",
    subject: content.subject,
    from: content.from,
    date: content.date,
    badges: content.badges,
    body: content.body,
    isHtml: content.isHtml,
    attachments: content.attachments.map((a) => ({
      name: a.name,
      type: a.type,
      base64: toBase64(a.data),
    })),
  };
}

// Opens the read dialog and, once it signals ready, posts the decrypted
// message to it in memory (chunked, since messageChild caps each frame at
// ~32KB). Resolves as soon as the dialog window is open — the taskpane
// does not wait for it to close. Rejects if displayDialogAsync fails.
function openDecryptedDialog(content: DecryptedContent): Promise<void> {
  const url = `${ADDIN_PUBLIC_URL}read-dialog.html`;
  const screenW = window.screen?.width || 1920;
  const screenH = window.screen?.height || 1080;
  const options: Office.DialogOptions = {
    width: pctOfScreen(900, screenW),
    height: pctOfScreen(800, screenH),
    displayInIframe: false,
    // Respect the same "skip the open-a-dialog confirmation" setting the
    // encrypt flow uses. Default (off) shows Office's prompt, which opens
    // reliably on every host; a blocked open falls back to the taskpane.
    promptBeforeOpen: !getAllowOptimisticDialog(),
  };

  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url, options, (asyncResult) => {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        const err = asyncResult.error;
        reject(err ? new Error(stringifyError(err)) : new Error("displayDialogAsync failed"));
        return;
      }
      const dialog = asyncResult.value;
      let sent = false;
      dialog.addEventHandler(
        Office.EventType.DialogMessageReceived,
        (arg: { message: string } | { error: number }) => {
          if ("error" in arg) return;
          let body: { type?: unknown };
          try {
            body = JSON.parse(arg.message) as { type?: unknown };
          } catch {
            return;
          }
          if (body.type === "ready" && !sent) {
            sent = true;
            const payload = buildDialogPayload(content);
            for (const c of chunkPayload(payload)) {
              dialog.messageChild(JSON.stringify(c));
            }
          }
        }
      );
      resolve();
    });
  });
}

// Compact taskpane state shown while the decrypted message lives in its
// own dialog window. Offers a button to re-open it (the plaintext is kept
// in memory on state.decrypted).
function showDialogOpenedView(): void {
  const text = byId<HTMLElement>("pg-decrypted-dialog-text");
  text.textContent = t("decryptedOpenedInWindow");

  const btn = byId<HTMLButtonElement>("pg-btn-reopen-decrypted");
  btn.textContent = t("decryptedReopen");
  // Replace listeners by cloning so re-renders don't stack handlers.
  const fresh = btn.cloneNode(true) as HTMLButtonElement;
  btn.replaceWith(fresh);
  fresh.addEventListener("click", () => void presentDecrypted());

  showView("decrypted_dialog");
}

// Fallback renderer: shows the decrypted message inline in the taskpane
// when the popup dialog could not be opened.
function renderInTaskpane(content: DecryptedContent): void {
  const subjectEl = byId<HTMLElement>("pg-decrypted-subject");
  subjectEl.textContent = content.subject;

  const metaEl = byId<HTMLElement>("pg-decrypted-meta");
  metaEl.textContent = [
    content.from && `${t("metaFrom")}: ${content.from}`,
    content.date && `${t("metaDate")}: ${content.date}`,
  ]
    .filter(Boolean)
    .join("  •  ");

  const badgesEl = byId<HTMLElement>("pg-decrypted-badges");
  badgesEl.innerHTML = "";
  if (content.badges.length > 0) {
    const label = document.createElement("span");
    label.textContent = `${t("notificationHeaderBadgesLabel")}: `;
    label.className = "pg-meta";
    badgesEl.appendChild(label);
    for (const value of content.badges) {
      const span = document.createElement("span");
      span.className = "pg-badge";
      span.textContent = value;
      badgesEl.appendChild(span);
    }
  }

  const iframe = byId<HTMLIFrameElement>("pg-decrypted-body");
  iframe.srcdoc = wrapHtml(content.body, content.isHtml);

  renderAttachments(content.attachments);

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
