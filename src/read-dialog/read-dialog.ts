// Read-mode decryption dialog. The taskpane decrypts the PostGuard
// envelope in memory and marshals the plaintext message here via
// Office.context.ui.messageChild, so a full email can be read in a roomy,
// normal-email-like window instead of the cramped taskpane (issue #72).
//
// The decrypted content never touches disk: it arrives as chunked
// in-memory messages, is rendered into the same sandboxed iframe +
// wrapHtml() used by the taskpane, and lives only for the lifetime of
// this dialog window. Closing the dialog drops it.

import { wrapHtml } from "../lib/render-body";
import { fromBase64 } from "../lib/encoding";
import { ChunkAssembler, isChunkMessage, ChunkMessage } from "../lib/dialog-chunk";
import { stringifyError } from "../lib/stringify-error";
import { t } from "../lib/i18n";

export interface DecryptedAttachmentPayload {
  name: string;
  type: string;
  base64: string;
}

// In-memory payload posted from the taskpane after a successful decrypt.
// `body` is the raw message body (HTML or plain text); the dialog runs it
// through the shared wrapHtml() so styling/contrast handling matches the
// taskpane exactly. No new sanitizer is introduced — rendering stays inside
// the sandboxed (no allow-scripts) iframe.
export interface DecryptedMessagePayload {
  type: "decrypted-message";
  subject: string;
  from: string;
  date: string;
  badges: string[];
  body: string;
  isHtml: boolean;
  attachments: DecryptedAttachmentPayload[];
}

const inboundChunks = new ChunkAssembler();

function log(msg: string): void {
  console.log(`[pg-read-dialog] ${msg}`);
}

function showError(message: string): void {
  const status = document.getElementById("pg-rd-status");
  if (status) status.hidden = true;
  const el = document.getElementById("pg-rd-error");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}

let attachmentObjectUrls: string[] = [];

function renderAttachments(attachments: DecryptedAttachmentPayload[]): void {
  // Revoke blobs from a previous render to free memory.
  for (const url of attachmentObjectUrls) URL.revokeObjectURL(url);
  attachmentObjectUrls = [];

  const host = document.getElementById("pg-rd-attachments");
  if (!host) return;
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
    const bytes = fromBase64(att.base64);
    const blob = new Blob([bytes as BlobPart], {
      type: att.type || "application/octet-stream",
    });
    const url = URL.createObjectURL(blob);
    attachmentObjectUrls.push(url);
    const li = document.createElement("li");
    const a = document.createElement("a");
    a.href = url;
    a.download = att.name;
    a.textContent = att.name;
    a.className = "pg-attachment-link";
    const size = document.createElement("span");
    size.className = "pg-meta";
    size.textContent = `  (${formatSize(bytes.byteLength)})`;
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

function render(msg: DecryptedMessagePayload): void {
  const status = document.getElementById("pg-rd-status");
  if (status) status.hidden = true;
  const content = document.getElementById("pg-rd-content");
  if (content) content.hidden = false;

  const subjectEl = document.getElementById("pg-rd-subject");
  if (subjectEl) subjectEl.textContent = msg.subject;

  const metaEl = document.getElementById("pg-rd-meta");
  if (metaEl) {
    metaEl.textContent = [
      msg.from && `${t("metaFrom")}: ${msg.from}`,
      msg.date && `${t("metaDate")}: ${msg.date}`,
    ]
      .filter(Boolean)
      .join("  •  ");
  }

  const badgesEl = document.getElementById("pg-rd-badges");
  if (badgesEl) {
    badgesEl.innerHTML = "";
    if (msg.badges.length > 0) {
      const label = document.createElement("span");
      label.textContent = `${t("notificationHeaderBadgesLabel")}: `;
      label.className = "pg-meta";
      badgesEl.appendChild(label);
      for (const value of msg.badges) {
        const span = document.createElement("span");
        span.className = "pg-badge";
        span.textContent = value;
        badgesEl.appendChild(span);
      }
    }
  }

  const iframe = document.getElementById("pg-rd-body") as HTMLIFrameElement | null;
  if (iframe) iframe.srcdoc = wrapHtml(msg.body, msg.isHtml);

  renderAttachments(msg.attachments);
}

function handlePayload(msg: { type?: unknown }): void {
  if (msg && msg.type === "decrypted-message") {
    render(msg as DecryptedMessagePayload);
  } else {
    log(`unknown payload type: ${String(msg?.type)}`);
  }
}

Office.onReady(() => {
  log("Office.onReady fired");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg: { message: string }) => {
      let payload: { type?: unknown };
      try {
        payload = JSON.parse(arg.message) as { type?: unknown };
      } catch (e) {
        log(`failed to parse parent message: ${stringifyError(e)}`);
        return;
      }
      if (isChunkMessage(payload)) {
        const reassembled = inboundChunks.ingest(payload as ChunkMessage);
        if (reassembled) handlePayload(reassembled as { type?: unknown });
        return;
      }
      handlePayload(payload);
    },
    (asyncResult) => {
      log(`addHandlerAsync status=${asyncResult.status}`);
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        showError(t("decryptionError"));
        return;
      }
      // Signal the taskpane that we're ready to receive the message.
      Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    }
  );
});
