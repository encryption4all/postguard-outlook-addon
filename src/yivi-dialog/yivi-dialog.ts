// Yivi-hosting dialog opened by the OnMessageSend handler for the
// one-click encrypt flow. The dialog runs in its own WebView2 with no
// access to Office.context.mailbox.item — the handler marshals the
// message data here via Office.context.ui.messageChild, the dialog
// runs pg-js + the Yivi QR widget, then posts the encrypted envelope
// back via messageParent.
//
// Payloads are chunked because messageChild/messageParent caps each
// frame at ~32KB and attachment bytes blow past that easily.

import { PostGuard, buildMime, UploadSessionExpiredError } from "@e4a/pg-js";
import { toBase64, fromBase64 } from "../lib/encoding";
import { PKG_URL, CRYPTIFY_URL, POSTGUARD_WEBSITE_URL } from "../lib/pkg-client";
import { ChunkAssembler, chunkPayload, isChunkMessage, ChunkMessage } from "../lib/dialog-chunk";
import { stringifyError } from "../lib/stringify-error";
import {
  recordPendingUpload,
  clearPendingUpload,
  probeAndClearPendingUpload,
} from "../lib/pending-upload";

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
  // Sender sign attributes built from the user's Settings prefills. Each
  // entry is either { t, v } (mandatory disclosure match — Yivi must
  // present a credential whose attribute equals `v`) or { t, optional: true }
  // (the Yivi app prompts, the user can skip).
  signAttributes?: { t: string; v?: string; optional?: boolean }[];
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

  // Forward the launchevent-built list verbatim — each entry is already
  // either { t, v } (mandatory match, from a Settings prefill) or
  // { t, optional: true } (no prefill). The previous code stripped `v`
  // and forced every entry back to optional, defeating the prefill flow.
  const signAttrs = (req.signAttributes ?? []).map((a) => ({
    t: a.t,
    ...(a.v !== undefined ? { v: a.v } : {}),
    ...(a.optional ? { optional: true } : {}),
  }));
  log(
    `sign attributes: ${
      signAttrs.map((a) => `${a.t}${a.v ? `=${a.v}` : a.optional ? ":optional" : ""}`).join(", ") ||
      "<none>"
    }`
  );

  const sealed = pg.encrypt({
    sign: pg.sign.yivi({
      element: "#yivi-web-form",
      senderEmail: req.senderEmail,
      includeSender: true,
      attributes: signAttrs.length ? signAttrs : undefined,
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
    onUploadInit: (info: { uuid: string; recoveryToken: string }) =>
      recordPendingUpload("local", info),
  } as never);
  clearPendingUpload("local");

  setSubtitle("Encrypting…");
  // pg-js 1.1.0+: envelope.attachment is null in tier 3 (the encrypted
  // payload was too large to ship as a local attachment; the body has
  // the Cryptify download link instead).
  let attBase64: string | null = null;
  if (envelope.attachment) {
    const attBytes = new Uint8Array(await envelope.attachment.arrayBuffer());
    attBase64 = toBase64(attBytes);
  }
  log(
    `tier=${envelope.tier} uploadUuid=${envelope.uploadUuid ?? "null"} attLen=${attBase64?.length ?? 0}`
  );

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
      // pg-js raises UploadSessionExpiredError when cryptify's structured
      // 404 says the upload session is gone (TTL expired, server restart,
      // unknown UUID, or wrong recovery_token). Surface that distinctly so
      // the Smart Alert can tell the user "start a new send" instead of a
      // generic "encryption failed" — see issue #82.
      const expired = err instanceof UploadSessionExpiredError;
      const message = expired
        ? "The upload session expired. Please start a new send."
        : stringifyError(err);
      log(`encryption failed${expired ? " (upload session expired)" : ""}: ${stringifyError(err)}`);
      showError(expired ? message : `Encryption failed: ${message}`);
      postChunkedToParent(
        expired
          ? { type: "encrypt-error", message, code: "upload_session_expired" }
          : { type: "encrypt-error", message }
      );
      showCompleted(message, true);
    }
  );
}

Office.onReady(() => {
  log("Office.onReady fired");

  // If a previous send left a recovery token behind, probe cryptify to
  // learn whether the session is still alive. pg-js 1.8.0 doesn't yet
  // accept a pre-resumed FileState back into createEnvelope, so the
  // probe is diagnostic: stale-session entries are dropped, live ones
  // are reported in the log and cleared (the user has to resend).
  void probeAndClearPendingUpload("local", CRYPTIFY_URL)
    .then((uploaded) => {
      if (uploaded !== null) log(`prior upload had ${uploaded} bytes on cryptify; entry cleared`);
    })
    .catch(() => undefined);

  // Show a hint in Safari pointing at the per-site popup setting.
  // Match Safari only — not WKWebView (Outlook for Mac), where no
  // such setting is reachable. Safari's UA includes "Safari/<ver>"
  // after AppleWebKit; WKWebView omits the Safari token. Chromium
  // browsers spoof AppleWebKit but include "Chrome" or "Edg".
  const ua = navigator.userAgent || "";
  const isSafari = /AppleWebKit/.test(ua) && /Safari\//.test(ua) && !/Chrome|Edg|OPR\//.test(ua);
  if (isSafari) {
    const tip = document.getElementById("pg-dlg-safari-tip");
    if (tip) tip.hidden = false;
  }

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
