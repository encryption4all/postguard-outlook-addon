// OnMessageSend handler. Runs in a separate WebView runtime from the
// taskpane, so it cannot read in-memory taskpane state. It uses x-
// prefixed internet headers set by the taskpane plus the attachment
// list to decide whether the message is allowed through.
//
// Behavior:
//  - encrypt-on-send not requested            → allow.
//  - requested + encrypted + recipients match → allow.
//  - requested + not yet encrypted            → open Yivi dialog, encrypt
//                                               in-line, apply result,
//                                               then allow.
//  - requested + encrypted + recipients drift → block (re-encrypt prompt).
//
// v1 of the one-click flow: text-only messages with email-only policy
// and email-only sign. Attachments and custom policy/sign require the
// manual taskpane "Encrypt & Send" flow until those are marshalled
// through to the dialog.

/* global Office */

import {
  ChunkAssembler,
  chunkPayload,
  isChunkMessage,
  ChunkMessage,
} from "../lib/dialog-chunk";
import { ADDIN_PUBLIC_URL } from "../lib/pkg-client";

const HEADER_ENCRYPT_ON_SEND = "x-pg-encrypt-on-send";
const HEADER_ENCRYPTED_RECIPIENTS = "x-pg-encrypted-recipients";
const HEADER_POSTGUARD = "x-postguard";
const POSTGUARD_VERSION = "0.1.0";
const POSTGUARD_ENCRYPTED_FILENAME = "postguard.encrypted";
const COMPOSE_BUTTON_ID = "postGuardComposeButton";
// Build the dialog URL from the add-in's public origin, injected at
// build time. window.location.href is unreliable here: New Outlook for
// Mac runs launchevent.js via the JSRuntime.Url override, where
// window.location resolves to an Office-internal URL rather than the
// add-in origin, and displayDialogAsync rejects with "An internal error
// has occurred."
const YIVI_DIALOG_URL = `${ADDIN_PUBLIC_URL}yivi-dialog.html`;

const NOT_ENCRYPTED_MESSAGE =
  "PostGuard is on but this message is not encrypted yet. " +
  "Open the PostGuard taskpane and click Encrypt & Send.";

const STALE_ENCRYPTION_MESSAGE =
  "PostGuard recipients or settings changed since the last encryption. " +
  "Open the PostGuard taskpane and click Re-encrypt & Send before sending.";

interface DialogMessage {
  type: string;
  [key: string]: unknown;
}

interface EncryptResult {
  subject: string;
  htmlBody: string;
  /** null in tier 3 — no local attachment to add (Cryptify-only flow). */
  attachmentBase64: string | null;
  tier: "tier1" | "tier2" | "tier3";
  uploadUuid: string | null;
}

function log(msg: string): void {
  // eslint-disable-next-line no-console
  console.log(`[pg-launchevent] ${msg}`);
}

function allowAfterTimeout(event: Office.AddinCommands.Event, ms = 270000): () => void {
  // 4½ min — gives the user time to find their phone and scan the QR.
  // Outlook's own Smart Alerts hard-cap is 5 min so we stay just under.
  const timer = setTimeout(() => {
    log(`fallback timeout (${ms}ms) reached; allowing the send`);
    try {
      event.completed({ allowEvent: true });
    } catch (e) {
      log(`fallback event.completed threw: ${String(e)}`);
    }
  }, ms);
  return () => clearTimeout(timer);
}

function block(event: Office.AddinCommands.Event, errorMessage: string): void {
  const opts: Office.SmartAlertsEventCompletedOptions = {
    allowEvent: false,
    errorMessage,
    commandId: COMPOSE_BUTTON_ID,
  };
  event.completed(opts);
}

function recipientsKey(addresses: Office.EmailAddressDetails[]): string {
  return addresses
    .map((a) => (a.emailAddress ?? "").toLowerCase().trim())
    .filter(Boolean)
    .sort()
    .join(",");
}

function getRecipientsAsync(
  recipients: Office.Recipients
): Promise<Office.EmailAddressDetails[]> {
  return new Promise((resolve) => {
    recipients.getAsync((res) =>
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : [])
    );
  });
}

function getSubjectAsync(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve, reject) => {
    item.subject.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

function setSubjectAsync(item: Office.MessageCompose, value: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.subject.setAsync(value, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function getBodyHtmlAsync(item: Office.MessageCompose): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

function setBodyHtmlAsync(item: Office.MessageCompose, value: string): Promise<void> {
  return new Promise((resolve, reject) => {
    item.body.setAsync(value, { coercionType: Office.CoercionType.Html }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function addBase64AttachmentAsync(
  item: Office.MessageCompose,
  filename: string,
  base64: string
): Promise<string> {
  return new Promise((resolve, reject) => {
    item.addFileAttachmentFromBase64Async(base64, filename, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded)
        resolve(res.value as unknown as string);
      else reject(res.error);
    });
  });
}

function getAttachmentContentAsync(
  item: Office.MessageCompose,
  attachmentId: string
): Promise<Office.AttachmentContent> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

function removeAttachmentAsync(
  item: Office.MessageCompose,
  attachmentId: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    item.removeAttachmentAsync(attachmentId, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function guessContentType(name: string): string {
  const ext = name.toLowerCase().split(".").pop() ?? "";
  const map: Record<string, string> = {
    pdf: "application/pdf",
    txt: "text/plain",
    csv: "text/csv",
    html: "text/html",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    png: "image/png",
    gif: "image/gif",
    zip: "application/zip",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  };
  return map[ext] ?? "application/octet-stream";
}

function setHeadersAsync(
  item: Office.MessageCompose,
  headers: Record<string, string>
): Promise<void> {
  return new Promise((resolve, reject) => {
    item.internetHeaders.setAsync(headers, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

function saveItemAsync(item: Office.MessageCompose): Promise<void> {
  return new Promise((resolve, reject) => {
    item.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

// Target physical size of the Yivi dialog. Just large enough for the
// QR widget (~250×280) plus title and Cancel button. We compute a
// screen-percentage from these at runtime because Office.displayDialog
// only accepts percentages — picking fixed percentages gives a tiny
// dialog on ultrawide monitors and an oversized one on laptops.
const YIVI_DIALOG_TARGET_WIDTH_PX = 300;
const YIVI_DIALOG_TARGET_HEIGHT_PX = 520;

// Flip to true to keep the Yivi dialog open after a successful encrypt
// (and after an encryption error) instead of auto-closing. Useful when
// debugging the dialog runtime — DevTools, log inspection, chunk
// reassembly, etc. Errors and the cancel path are unaffected; cancel
// always closes itself.
const DEBUG_KEEP_DIALOG_OPEN = false;

// Floor the dialog size at 30% of the screen. Office.js docs claim the
// valid range is 1–99, but in practice Outlook hosts (Web on ultrawide
// monitors, New Outlook for Mac) reject very small percentages with no
// further detail — Web returns `code=12011 BlockedNavigation`, Mac
// surfaces it as the generic E_FAIL. 30% comfortably clears whatever
// the actual minimum is across hosts and still fits a 250-300px QR.
const MIN_DIALOG_PCT = 30;

function pctOfScreen(targetPx: number, screenPx: number): number {
  const pct = Math.ceil((targetPx / screenPx) * 100);
  return Math.min(99, Math.max(MIN_DIALOG_PCT, pct));
}

// Opens the Yivi dialog with an encrypt-request payload and waits for
// the dialog to post the encrypted result back. Resolves with the
// envelope; rejects on error or user cancel.
function runEncryptDialog(payload: DialogMessage): Promise<EncryptResult> {
  return new Promise((resolve, reject) => {
    // window.screen.* is in CSS pixels (matching what Office's
    // percentage interprets). Falls back to a safe 1920×1080 if Office's
    // launchevent runtime ever surfaces an empty screen object.
    const screenW = window.screen?.width || 1920;
    const screenH = window.screen?.height || 1080;
    const widthPct = pctOfScreen(YIVI_DIALOG_TARGET_WIDTH_PX, screenW);
    const heightPct = pctOfScreen(YIVI_DIALOG_TARGET_HEIGHT_PX, screenH);
    log(`dialog size: target ${YIVI_DIALOG_TARGET_WIDTH_PX}×${YIVI_DIALOG_TARGET_HEIGHT_PX}px on ${screenW}×${screenH} screen → ${widthPct}%×${heightPct}%`);

    Office.context.ui.displayDialogAsync(
      YIVI_DIALOG_URL,
      // promptBeforeOpen: false suppresses the "PostGuard is opening
      // another window" confirmation. Honored because the dialog URL is
      // on the same origin as the add-in's source location. Requires
      // Mailbox 1.9 (we require 1.12 in VersionOverridesV1_1).
      { height: heightPct, width: widthPct, displayInIframe: false, promptBeforeOpen: false },
      (asyncResult) => {
        log(`displayDialogAsync status=${asyncResult.status}`);
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(`displayDialogAsync failed: ${asyncResult.error?.message}`));
          return;
        }
        const dialog = asyncResult.value;
        const inbound = new ChunkAssembler();
        let settled = false;
        // Auto-close on success/error so the user isn't left with a
        // stale "Encrypted and sent. You can close this window." dialog
        // after the Send has been released — flip DEBUG_KEEP_DIALOG_OPEN
        // to opt out when DevTools/log inspection is needed. Cancel
        // closes itself from the dialog (window.close on the button).
        const closeDialog = (): void => {
          if (DEBUG_KEEP_DIALOG_OPEN) return;
          try {
            dialog.close();
          } catch (e) {
            log(`dialog.close failed: ${String(e)}`);
          }
        };
        const settle = (cb: () => void): void => {
          if (settled) return;
          settled = true;
          cb();
        };

        const dispatch = (body: DialogMessage): void => {
          log(`dialog → handler: ${body.type}`);
          switch (body.type) {
            case "ready": {
              const chunks = chunkPayload(payload);
              log(`sending ${chunks.length} chunk(s) to dialog`);
              for (const c of chunks) {
                dialog.messageChild(JSON.stringify(c));
              }
              break;
            }
            case "encrypt-result":
              settle(() => {
                closeDialog();
                resolve(body as unknown as EncryptResult);
              });
              break;
            case "encrypt-error":
              settle(() => {
                closeDialog();
                reject(new Error(String(body.message ?? "Encryption failed")));
              });
              break;
            case "cancelled":
              settle(() => reject(new Error("Cancelled in dialog")));
              break;
            default:
              log(`unhandled dialog message: ${body.type}`);
          }
        };

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string } | { error: number }) => {
            if ("error" in arg) {
              log(`dialog message error: ${arg.error}`);
              settle(() => reject(new Error(`Dialog error ${arg.error}`)));
              return;
            }
            let body: DialogMessage;
            try {
              body = JSON.parse(arg.message) as DialogMessage;
            } catch {
              log(`could not parse dialog message: ${arg.message}`);
              return;
            }
            if (isChunkMessage(body)) {
              const reassembled = inbound.ingest(body as ChunkMessage);
              if (reassembled) dispatch(reassembled as DialogMessage);
              return;
            }
            dispatch(body);
          }
        );

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          log(`dialog event: ${JSON.stringify(arg)}`);
          if ("error" in arg && arg.error === 12006) {
            settle(() => reject(new Error("Dialog closed by user")));
          }
        });
      }
    );
  });
}

async function readUserAttachments(
  item: Office.MessageCompose,
  attachments: Office.AttachmentDetailsCompose[]
): Promise<{ name: string; type: string; base64: string }[]> {
  const out: { name: string; type: string; base64: string }[] = [];
  for (const a of attachments) {
    // Skip cloud attachments — Office.js can't read their bytes.
    if (a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud) {
      log(`skipping cloud attachment: ${a.name}`);
      continue;
    }
    try {
      const content = await getAttachmentContentAsync(item, a.id);
      const base64Len = content.content?.length ?? 0;
      log(
        `attachment "${a.name}" format=${content.format} ` +
          `base64Len=${base64Len} declaredSize=${a.size ?? "?"}`
      );
      // Tenant DLP can scrub attachment bytes (e.g. blocked extensions like
      // .exe) while still reporting metadata. Detect: declared size > 0 but
      // returned content is empty. We refuse rather than silently encrypt
      // a 0-byte attachment.
      if ((a.size ?? 0) > 0 && base64Len === 0) {
        throw new Error(
          `Outlook returned no content for attachment "${a.name}" — ` +
            `your tenant likely blocks this file type. ` +
            `Remove the attachment or zip it with a different extension.`
        );
      }
      if (content.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
        out.push({ name: a.name, type: guessContentType(a.name), base64: content.content });
      } else {
        log(`unsupported attachment format for ${a.name}: ${content.format}`);
      }
    } catch (e) {
      log(`failed to read attachment ${a.name}: ${String(e)}`);
      throw e;
    }
  }
  return out;
}

async function encryptAndApply(
  event: Office.AddinCommands.Event,
  item: Office.MessageCompose,
  to: Office.EmailAddressDetails[],
  cc: Office.EmailAddressDetails[],
  userAttachments: Office.AttachmentDetailsCompose[]
): Promise<void> {
  const senderEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  const subject = await getSubjectAsync(item);
  const htmlBody = await getBodyHtmlAsync(item);
  const attachments = await readUserAttachments(item, userAttachments);

  const result = await runEncryptDialog({
    type: "encrypt-request",
    senderEmail,
    to: to.map((r) => r.emailAddress.toLowerCase()),
    cc: cc.map((r) => r.emailAddress.toLowerCase()),
    subject,
    htmlBody,
    attachments,
  });

  await setSubjectAsync(item, result.subject);
  await setBodyHtmlAsync(item, result.htmlBody);
  // Remove the original plaintext attachments now that they're inside the
  // encrypted envelope. Best-effort: a cloud attachment we couldn't read
  // would still be sent in the clear, so we leave it alone.
  for (const a of userAttachments) {
    if (a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud) continue;
    try {
      await removeAttachmentAsync(item, a.id);
    } catch (e) {
      log(`failed to remove original attachment ${a.name}: ${String(e)}`);
    }
  }
  // Tier 1/2: include the encrypted bytes locally as postguard.encrypted.
  // Tier 3: pg-js gave us no attachment (too large) — recipients use the
  // Cryptify link in the body to fetch and decrypt.
  if (result.attachmentBase64) {
    await addBase64AttachmentAsync(item, POSTGUARD_ENCRYPTED_FILENAME, result.attachmentBase64);
  } else {
    log(`tier ${result.tier}: skipping local attachment, recipients fetch via uuid=${result.uploadUuid}`);
  }
  await setHeadersAsync(item, {
    [HEADER_ENCRYPTED_RECIPIENTS]: recipientsKey([...to, ...cc]),
    [HEADER_POSTGUARD]: POSTGUARD_VERSION,
  });
  await saveItemAsync(item);
}

function onMessageSendHandler(event: Office.AddinCommands.Event): void {
  log("onMessageSendHandler invoked");
  const cancelTimeout = allowAfterTimeout(event);

  const item = Office.context.mailbox.item as Office.MessageCompose;
  if (!item) {
    log("no mailbox item; allowing");
    cancelTimeout();
    event.completed({ allowEvent: true });
    return;
  }

  item.internetHeaders.getAsync(
    [HEADER_ENCRYPT_ON_SEND, HEADER_ENCRYPTED_RECIPIENTS],
    (hdrRes) => {
      log(`internetHeaders.getAsync status=${hdrRes.status}`);
      if (hdrRes.status !== Office.AsyncResultStatus.Succeeded) {
        cancelTimeout();
        event.completed({ allowEvent: true });
        return;
      }

      const encryptRequested = hdrRes.value[HEADER_ENCRYPT_ON_SEND] === "true";
      log(`encryptRequested=${encryptRequested}`);
      if (!encryptRequested) {
        cancelTimeout();
        event.completed({ allowEvent: true });
        return;
      }

      const stampedRecipients = hdrRes.value[HEADER_ENCRYPTED_RECIPIENTS] ?? "";

      item.getAttachmentsAsync(async (attRes) => {
        log(`getAttachmentsAsync status=${attRes.status}`);
        const attachments =
          attRes.status === Office.AsyncResultStatus.Succeeded ? attRes.value : [];
        const alreadyEncrypted = attachments.some(
          (a) => a.name?.toLowerCase() === POSTGUARD_ENCRYPTED_FILENAME
        );
        log(`alreadyEncrypted=${alreadyEncrypted} (${attachments.length} attachments)`);

        const [to, cc] = await Promise.all([
          getRecipientsAsync(item.to),
          getRecipientsAsync(item.cc),
        ]);

        if (!alreadyEncrypted) {
          if (to.length + cc.length === 0) {
            cancelTimeout();
            block(event, "Add at least one recipient before sending.");
            return;
          }

          try {
            await encryptAndApply(event, item, to, cc, attachments);
            cancelTimeout();
            event.completed({ allowEvent: true });
          } catch (e) {
            cancelTimeout();
            const msg = e instanceof Error ? e.message : String(e);
            block(event, `Encryption failed: ${msg}`);
          }
          return;
        }

        // Verify the encryption matches the message's current To+Cc list.
        const currentKey = recipientsKey([...to, ...cc]);
        const stale = stampedRecipients === "" || currentKey !== stampedRecipients;
        log(`stamped=${stampedRecipients || "<empty>"} current=${currentKey} stale=${stale}`);

        cancelTimeout();
        if (stale) {
          block(event, STALE_ENCRYPTION_MESSAGE);
          return;
        }
        event.completed({ allowEvent: true });
      });
    }
  );
}

log("script loaded");
Office.onReady((info) => {
  log(`Office.onReady fired; host=${info?.host} platform=${info?.platform}`);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  log("handler associated");
});
