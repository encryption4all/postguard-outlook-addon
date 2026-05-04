// Yivi-hosting dialog opened by the OnMessageSend handler for the
// one-click encrypt flow. The dialog runs in its own WebView2 with no
// access to Office.context.mailbox.item — the handler marshals the
// message data here via Office.context.ui.messageChild, the dialog
// runs pg-js + the Yivi QR widget, then posts the encrypted envelope
// back via messageParent.
//
// Payloads are chunked because messageChild/messageParent caps each
// frame at ~32KB and attachment bytes blow past that easily.

import { PostGuard, buildMime } from "@e4a/pg-js";
import { toBase64, fromBase64 } from "../lib/encoding";
import { PKG_URL, CRYPTIFY_URL, POSTGUARD_WEBSITE_URL } from "../lib/pkg-client";
import {
  ChunkAssembler,
  chunkPayload,
  isChunkMessage,
  ChunkMessage,
} from "../lib/dialog-chunk";

/* global Office */

const ADDIN_VERSION = "0.1.0";

interface AttachmentPayload {
  name: string;
  type: string;
  base64: string;
}

interface EncryptRequest {
  type: "encrypt-request";
  senderEmail: string;
  to: string[];
  cc: string[];
  subject: string;
  htmlBody: string;
  attachments: AttachmentPayload[];
}

interface EncryptResult {
  type: "encrypt-result";
  subject: string;
  htmlBody: string;
  /** null in tier 3 — no local attachment, the body's Cryptify link
   *  carries the ciphertext for recipients. */
  attachmentBase64: string | null;
  tier: "tier1" | "tier2" | "tier3";
  uploadUuid: string | null;
}

interface DialogMessage {
  type: string;
  [key: string]: unknown;
}

const inboundChunks = new ChunkAssembler();

function log(msg: string): void {
  // eslint-disable-next-line no-console
  console.log(`[pg-dialog] ${msg}`);
}

function setSubtitle(text: string): void {
  const el = document.getElementById("pg-dlg-subtitle");
  if (el) el.textContent = text;
}

function setTitle(text: string): void {
  const el = document.getElementById("pg-dlg-title");
  if (el) el.textContent = text;
}

function showError(message: string): void {
  const el = document.getElementById("pg-dlg-error");
  if (!el) return;
  el.textContent = message;
  el.hidden = false;
}

// Switch the dialog into "completed" mode: hide the Yivi widget area
// and the Cancel button, show a Close button, and stop auto-closing.
// The Send is already released (the handler applied the result and
// called event.completed) — this just lets the user read any logs in
// DevTools before dismissing the dialog.
function showCompleted(message: string, isError = false): void {
  const yiviHost = document.getElementById("yivi-web-form");
  if (yiviHost) yiviHost.hidden = true;
  const cancelBtn = document.getElementById("pg-dlg-cancel");
  if (cancelBtn) cancelBtn.hidden = true;
  const closeBtn = document.getElementById("pg-dlg-close");
  if (closeBtn) closeBtn.hidden = false;
  setTitle(isError ? "Encryption failed" : "Done");
  setSubtitle(message);
}

function postChunkedToParent(payload: DialogMessage): void {
  const chunks = chunkPayload(payload);
  log(`posting ${chunks.length} chunk(s) to parent`);
  for (const c of chunks) {
    Office.context.ui.messageParent(JSON.stringify(c));
  }
}

async function runEncryption(req: EncryptRequest): Promise<EncryptResult> {
  setTitle("Sign your message");
  setSubtitle("Scan the QR code with the Yivi app to sign and send.");

  for (const a of req.attachments) {
    log(`received attachment "${a.name}" type=${a.type} base64Len=${a.base64.length}`);
  }
  const attachmentsForMime = req.attachments.map((a) => ({
    name: a.name,
    type: a.type,
    data: fromBase64(a.base64).buffer as ArrayBuffer,
  }));

  const mime = (await buildMime({
    from: req.senderEmail,
    to: req.to,
    cc: req.cc,
    subject: req.subject,
    htmlBody: req.htmlBody,
    date: new Date(),
    attachments: attachmentsForMime,
  } as never)) as Uint8Array;

  const pg = new PostGuard({
    pkgUrl: PKG_URL,
    cryptifyUrl: CRYPTIFY_URL,
    headers: {
      "X-PostGuard-Client-Version": `Outlook,1.0,pg4outlook,${ADDIN_VERSION}`,
    },
  } as never);

  const recipients = [...req.to, ...req.cc].map((email) =>
    (pg as never as { recipient: { email: (e: string) => unknown } }).recipient.email(email)
  );

  const sealed = pg.encrypt({
    sign: pg.sign.yivi({
      element: "#yivi-web-form",
      senderEmail: req.senderEmail,
    } as never),
    recipients,
    data: mime,
  } as never);

  // pg-js 1.2.0+: the Cryptify upload is silent by default — no
  // recipient notification is sent. The user's message is delivered
  // from their own email account, and the Cryptify upload provides
  // the in-body download link without producing a duplicate mail. We
  // therefore let createEnvelope upload for tier 2 and tier 3 alike.
  const envelope = await pg.email.createEnvelope({
    sealed,
    from: req.senderEmail,
    websiteUrl: POSTGUARD_WEBSITE_URL,
  } as never);

  setSubtitle("Encrypting…");
  // pg-js 1.1.0+: envelope.attachment is null in tier 3 (the encrypted
  // payload was too large to ship as a local attachment; the body has
  // the Cryptify download link instead).
  let attBase64: string | null = null;
  if (envelope.attachment) {
    const attBytes = new Uint8Array(await envelope.attachment.arrayBuffer());
    attBase64 = toBase64(attBytes);
  }
  log(`tier=${envelope.tier} uploadUuid=${envelope.uploadUuid ?? "null"} attLen=${attBase64?.length ?? 0}`);

  return {
    type: "encrypt-result",
    subject: envelope.subject,
    htmlBody: envelope.htmlBody,
    attachmentBase64: attBase64,
    tier: envelope.tier,
    uploadUuid: envelope.uploadUuid,
  };
}

function handlePayload(msg: DialogMessage): void {
  log(`payload type=${msg.type}`);
  if (msg.type !== "encrypt-request") {
    log(`unknown payload type: ${msg.type}`);
    return;
  }
  const req = msg as unknown as EncryptRequest;
  setSubtitle(
    `Building encrypted message (${req.attachments.length} attachment${req.attachments.length === 1 ? "" : "s"})…`
  );
  void runEncryption(req).then(
    (result) => {
      log("encryption complete; posting result");
      postChunkedToParent(result as unknown as DialogMessage);
      showCompleted("Encrypted and sent. You can close this window.");
    },
    (err) => {
      const message = err instanceof Error ? err.message : String(err);
      log(`encryption failed: ${message}`);
      showError(`Encryption failed: ${message}`);
      postChunkedToParent({ type: "encrypt-error", message });
      showCompleted(message, true);
    }
  );
}

Office.onReady(() => {
  log("Office.onReady fired");

  const cancelBtn = document.getElementById("pg-dlg-cancel") as HTMLButtonElement | null;
  if (cancelBtn) {
    cancelBtn.addEventListener("click", () => {
      postChunkedToParent({ type: "cancelled" });
      window.close();
    });
  }

  const closeBtn = document.getElementById("pg-dlg-close") as HTMLButtonElement | null;
  if (closeBtn) {
    closeBtn.addEventListener("click", () => {
      window.close();
    });
  }

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg: { message: string }) => {
      let payload: DialogMessage;
      try {
        payload = JSON.parse(arg.message) as DialogMessage;
      } catch (e) {
        log(`failed to parse parent message: ${String(e)}`);
        return;
      }
      if (isChunkMessage(payload)) {
        const reassembled = inboundChunks.ingest(payload as ChunkMessage);
        if (reassembled) handlePayload(reassembled as DialogMessage);
        return;
      }
      handlePayload(payload);
    },
    (asyncResult) => {
      log(`addHandlerAsync status=${asyncResult.status}`);
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        showError("Could not register message handler. Please retry.");
        return;
      }
      // ready is small, send unchunked.
      Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    }
  );
});
