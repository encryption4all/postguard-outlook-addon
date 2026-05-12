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

import { ChunkAssembler, chunkPayload, isChunkMessage, ChunkMessage } from "../lib/dialog-chunk";
import { ADDIN_PUBLIC_URL } from "../lib/pkg-client";
import {
  buildSignAttributes,
  getAllowOptimisticDialog,
  getEncryptionEnabled,
} from "../lib/settings";
import { stringifyError } from "../lib/stringify-error";
import { t } from "../lib/i18n";

const ENCRYPTION_STATUS_NOTIFICATION_KEY = "postguard-encryption-status";

// Per-draft "is this email going to be encrypted on send" flag. Written
// by the taskpane's compose toggle and seeded from the mailbox-wide
// default by OnNewMessageCompose, so by the time OnMessageSend fires
// the header reflects the user's explicit intent for *this* message.
// The OnMessageSend handler reads only this header — never the global
// setting — so a PostGuard outage can never block an unencrypted send.
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

const STALE_ENCRYPTION_MESSAGE =
  "PostGuard recipients or settings changed since the last encryption. " +
  "Open the PostGuard taskpane and click Re-encrypt & Send before sending.";

const MAC_NOT_SUPPORTED_MESSAGE =
  "PostGuard's one-click encrypt-on-send is not supported on Outlook for Mac. " +
  "Open the PostGuard taskpane (PostGuard button in the toolbar) and click " +
  "Encrypt & Send to encrypt and send this message.";

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
  console.log(`[pg-launchevent] ${msg}`);
}

// Encryption-path watchdog. Once the user opts into encrypting this
// message (x-pg-encrypt-on-send=true), we must never release the send
// in cleartext — silently sending an unencrypted email that the user
// asked to be encrypted is the worst possible failure mode. So if the
// encryption flow doesn't complete in time, block the send with a
// clear error and let the user retry.
//
// 4½ min: Outlook's Smart Alerts hard-cap is 5 min, so we stay just
// under. The user has that long to find their phone and scan the QR.
function blockAfterTimeout(onFire: () => void, ms = 270000): () => void {
  const timer = setTimeout(() => {
    log(`fallback timeout (${ms}ms) reached; blocking the send`);
    onFire();
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

function getRecipientsAsync(recipients: Office.Recipients): Promise<Office.EmailAddressDetails[]> {
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

function removeAttachmentAsync(item: Office.MessageCompose, attachmentId: string): Promise<void> {
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

// Target physical size of the Yivi dialog. Sized to fit the QR widget
// (~290px wide) with comfortable margins, plus title, optional Safari
// hint and Cancel button. We compute a screen-percentage from these at
// runtime because Office.displayDialog only accepts percentages —
// picking fixed percentages gives a tiny dialog on ultrawide monitors
// and an oversized one on laptops.
const YIVI_DIALOG_TARGET_WIDTH_PX = 460;
const YIVI_DIALOG_TARGET_HEIGHT_PX = 640;

// Flip to true to keep the Yivi dialog open after a successful encrypt
// (and after an encryption error) instead of auto-closing. Useful when
// debugging the dialog runtime — DevTools, log inspection, chunk
// reassembly, etc. Errors and the cancel path are unaffected; cancel
// always closes itself.
const DEBUG_KEEP_DIALOG_OPEN = false;

function pctOfScreen(targetPx: number, screenPx: number): number {
  // displayDialogAsync clamps to [1, 99]. Round up so we don't drop
  // below the QR's minimum useful size on huge monitors.
  const pct = Math.ceil((targetPx / screenPx) * 100);
  return Math.min(99, Math.max(1, pct));
}

// Promise wrapper around displayDialogAsync. Resolves with the dialog
// handle on success, rejects with the Office error otherwise.
function openDialogAsync(url: string, options: Office.DialogOptions): Promise<Office.Dialog> {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(url, options, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value);
      } else {
        const err = asyncResult.error;
        reject(err ? new Error(stringifyError(err)) : new Error("displayDialogAsync failed"));
      }
    });
  });
}

// Opens the Yivi dialog with an encrypt-request payload and waits for
// the dialog to post the encrypted result back. Resolves with the
// envelope; rejects on error or user cancel.
async function runEncryptDialog(payload: DialogMessage): Promise<EncryptResult> {
  // window.screen.* is in CSS pixels (matching what Office's percentage
  // interprets). Falls back to a safe 1920×1080 if the launchevent
  // runtime ever surfaces an empty screen object.
  const screenW = window.screen?.width || 1920;
  const screenH = window.screen?.height || 1080;
  const widthPct = pctOfScreen(YIVI_DIALOG_TARGET_WIDTH_PX, screenW);
  const heightPct = pctOfScreen(YIVI_DIALOG_TARGET_HEIGHT_PX, screenH);
  log(
    `dialog size: target ${YIVI_DIALOG_TARGET_WIDTH_PX}×${YIVI_DIALOG_TARGET_HEIGHT_PX}px on ${screenW}×${screenH} screen → ${widthPct}%×${heightPct}%`
  );

  const baseOptions: Office.DialogOptions = {
    height: heightPct,
    width: widthPct,
    displayInIframe: false,
  };

  // Default: open with Office's "PostGuard wants to open a dialog"
  // confirmation. The user's click on Allow is itself a fresh user
  // gesture, so the popup opens reliably on every host (including
  // Safari without site-level popup permission). Power users who have
  // permanently allowed pop-ups for the add-in's origin can flip the
  // Settings toggle (taskpane → gear → Skip the "open a dialog"
  // confirmation) to opt into a single-attempt optimistic open. If
  // that attempt is still blocked we recover by re-trying with the
  // prompt so the send isn't lost.
  const allowOptimistic = getAllowOptimisticDialog();
  log(`displayDialogAsync: promptBeforeOpen=${!allowOptimistic} (optimistic=${allowOptimistic})`);
  let dialog: Office.Dialog;
  try {
    dialog = await openDialogAsync(YIVI_DIALOG_URL, {
      ...baseOptions,
      promptBeforeOpen: !allowOptimistic,
    });
    log(allowOptimistic ? "dialog opened (no prompt)" : "dialog opened (after prompt)");
  } catch (e) {
    if (!allowOptimistic) throw e;
    log(`optimistic attempt failed (${stringifyError(e)}); retrying with promptBeforeOpen=true`);
    dialog = await openDialogAsync(YIVI_DIALOG_URL, {
      ...baseOptions,
      promptBeforeOpen: true,
    });
    log("dialog opened (after prompt fallback)");
  }

  return new Promise((resolve, reject) => {
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
        log(`dialog.close failed: ${stringifyError(e)}`);
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
            const raw = body.message;
            const text =
              typeof raw === "string"
                ? raw
                : raw === undefined
                  ? "Encryption failed"
                  : stringifyError(raw);
            reject(new Error(text));
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
      log(`failed to read attachment ${a.name}: ${stringifyError(e)}`);
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

  // Sender attributes come from per-mailbox roaming settings (configured
  // in the taskpane Settings view). Roaming settings are available in the
  // launchevent runtime, so we don't need a per-draft internet header.
  const signAttributes = buildSignAttributes();
  log(
    `signAttributes=${signAttributes
      .map((a) => `${a.t}${a.v ? `=${a.v}` : a.optional ? ":optional" : ""}`)
      .join(", ")}`
  );

  const result = await runEncryptDialog({
    type: "encrypt-request",
    senderEmail,
    to: to.map((r) => r.emailAddress.toLowerCase()),
    cc: cc.map((r) => r.emailAddress.toLowerCase()),
    subject,
    htmlBody,
    attachments,
    signAttributes,
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
      log(`failed to remove original attachment ${a.name}: ${stringifyError(e)}`);
    }
  }
  // Tier 1/2: include the encrypted bytes locally as postguard.encrypted.
  // Tier 3: pg-js gave us no attachment (too large) — recipients use the
  // Cryptify link in the body to fetch and decrypt.
  if (result.attachmentBase64) {
    await addBase64AttachmentAsync(item, POSTGUARD_ENCRYPTED_FILENAME, result.attachmentBase64);
  } else {
    log(
      `tier ${result.tier}: skipping local attachment, recipients fetch via uuid=${result.uploadUuid}`
    );
  }
  await setHeadersAsync(item, {
    [HEADER_ENCRYPTED_RECIPIENTS]: recipientsKey([...to, ...cc]),
    [HEADER_POSTGUARD]: POSTGUARD_VERSION,
  });
  await saveItemAsync(item);
}

function onMessageSendHandler(event: Office.AddinCommands.Event): void {
  // Two completion modes:
  //   - release  → event.completed({ allowEvent: true }).  Email goes
  //                out as-is (no PostGuard involvement). Use for the
  //                "off" or "indeterminate" paths so a broken add-in
  //                can never stop an unencrypted send from happening.
  //   - blockSend → event.completed({ allowEvent: false, errorMessage }).
  //                Outlook shows a Smart Alert and refuses the send.
  //                Use whenever we have committed to encrypting and
  //                something then went wrong — silently sending a
  //                "supposed to be encrypted" email in plaintext is
  //                the failure mode we never accept.
  //
  // `committedToEncrypt` flips to true the instant we confirm the
  // header is "true", so any error after that point routes to
  // blockSend instead of releaseSend.
  let settled = false;
  let committedToEncrypt = false;

  const releaseSend = (reason: string): void => {
    if (settled) return;
    settled = true;
    log(`releasing send: ${reason}`);
    try {
      event.completed({ allowEvent: true });
    } catch (e) {
      log(`event.completed threw on release: ${stringifyError(e)}`);
    }
  };

  const blockSend = (errorMessage: string): void => {
    if (settled) return;
    settled = true;
    log(`blocking send: ${errorMessage}`);
    try {
      block(event, errorMessage);
    } catch (e) {
      log(`block threw: ${stringifyError(e)}`);
    }
  };

  const onFailure = (reason: string): void => {
    if (committedToEncrypt) {
      blockSend(`PostGuard encryption failed: ${reason}`);
    } else {
      releaseSend(reason);
    }
  };

  try {
    log("onMessageSendHandler invoked");

    const item = Office.context.mailbox.item as Office.MessageCompose | undefined;
    if (!item || !item.internetHeaders) {
      releaseSend("no compose item / no internetHeaders");
      return;
    }

    // Single source of truth at send time: the per-draft header. The
    // header is seeded from the mailbox-wide default by
    // OnNewMessageCompose and updated by the compose toggle. We only
    // run the encryption path when the header says "true" — anything
    // else (false, absent, read failure) releases the send so a
    // PostGuard outage cannot block an unencrypted email from going
    // out. Once we *have* seen "true", encryption is non-negotiable.
    item.internetHeaders.getAsync(
      [HEADER_ENCRYPT_ON_SEND, HEADER_ENCRYPTED_RECIPIENTS],
      (hdrRes) => {
        if (settled) return;

        try {
          if (hdrRes.status !== Office.AsyncResultStatus.Succeeded) {
            releaseSend(`header read failed (status=${hdrRes.status})`);
            return;
          }

          const encryptHeader = hdrRes.value[HEADER_ENCRYPT_ON_SEND];
          if (encryptHeader !== "true") {
            releaseSend(`x-pg-encrypt-on-send=${encryptHeader ?? "<absent>"}`);
            return;
          }

          // From this point on, the user has explicitly asked for this
          // message to be encrypted. Any error must block the send.
          committedToEncrypt = true;
          const cancelTimeout = blockAfterTimeout(() =>
            blockSend(
              "PostGuard encryption did not finish in time. The message was NOT sent. " +
                "Try again, or turn PostGuard off in the taskpane to send this message unencrypted."
            )
          );
          const stampedRecipients = hdrRes.value[HEADER_ENCRYPTED_RECIPIENTS] ?? "";

          item.getAttachmentsAsync(async (attRes) => {
            // Inner guard: anything that throws (or any async rejection
            // we forgot to handle) inside this async callback must
            // route to blockSend, not bubble up as an unhandled rejection
            // that lets the Smart Alert timer eventually release the
            // send. We're past the committedToEncrypt point.
            try {
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
                  blockSend("Add at least one recipient before sending.");
                  return;
                }

                // Outlook for Mac (native WKWebView) rejects displayDialogAsync
                // from the launchevent runtime with E_FAIL regardless of options
                // or sizing. Tracked upstream at office-js#6677; related stale
                // reports are #3138, #3085, and #5681. Until Microsoft restores
                // working dialog support, deflect Mac users to the manual
                // taskpane "Encrypt & Send" button (which uses the dialog API
                // from the taskpane runtime, where it works). Note: this only
                // fires when the message is *not* already encrypted — once the
                // user has clicked Encrypt & Send in the taskpane and we see
                // the postguard.encrypted attachment, we fall through to the
                // standard allow-send path. Remove this branch when 6677 ships.
                if (Office.context.platform === Office.PlatformType.Mac) {
                  log("Outlook for Mac detected; deferring to taskpane flow");
                  cancelTimeout();
                  blockSend(MAC_NOT_SUPPORTED_MESSAGE);
                  return;
                }

                try {
                  await encryptAndApply(event, item, to, cc, attachments);
                  cancelTimeout();
                  if (!settled) {
                    settled = true;
                    event.completed({ allowEvent: true });
                  }
                } catch (e) {
                  cancelTimeout();
                  blockSend(`PostGuard encryption failed: ${stringifyError(e)}`);
                }
                return;
              }

              // Verify the encryption matches the message's current To+Cc list.
              const currentKey = recipientsKey([...to, ...cc]);
              const stale = stampedRecipients === "" || currentKey !== stampedRecipients;
              log(`stamped=${stampedRecipients || "<empty>"} current=${currentKey} stale=${stale}`);

              cancelTimeout();
              if (stale) {
                blockSend(STALE_ENCRYPTION_MESSAGE);
                return;
              }
              if (!settled) {
                settled = true;
                event.completed({ allowEvent: true });
              }
            } catch (e) {
              cancelTimeout();
              blockSend(`PostGuard encryption failed: ${stringifyError(e)}`);
            }
          });
        } catch (e) {
          onFailure(`unexpected error after header read: ${stringifyError(e)}`);
        }
      }
    );
  } catch (e) {
    onFailure(`unexpected error before main flow: ${stringifyError(e)}`);
  }
}

// OnNewMessageCompose handler. Fires when the user opens a new message,
// reply, or forward — independent of whether the user has clicked the
// PostGuard ribbon button. Two jobs:
//   1. Seed the per-draft x-pg-encrypt-on-send header from the
//      mailbox-wide default if it isn't already set. This makes the
//      header the single source of truth for the OnMessageSend handler
//      while still letting the user pick a default in Settings.
//   2. Set the persistent in-message "PostGuard is on/off" banner so
//      the user always sees the current state of this specific draft,
//      even before opening the taskpane.
function onNewMessageComposeHandler(event: Office.AddinCommands.Event): void {
  log("onNewMessageComposeHandler invoked");
  const item = Office.context.mailbox.item as Office.MessageCompose | undefined;
  if (!item || !item.notificationMessages || !item.internetHeaders) {
    log("no compose item / notificationMessages / internetHeaders; completing");
    event.completed();
    return;
  }

  let globalDefault = false;
  try {
    globalDefault = getEncryptionEnabled();
  } catch (e) {
    log(`failed to read encryption setting; assuming off: ${stringifyError(e)}`);
  }

  const applyBanner = (encryptOn: boolean): void => {
    const messageText = encryptOn
      ? t("composeEncryptionOnBanner")
      : t("composeEncryptionOffBanner");
    const details: Office.NotificationMessageDetails = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: messageText.slice(0, 150),
      icon: "Icon.16x16",
      persistent: true,
    };
    // Remove first, then add: replaceAsync silently no-ops on the same
    // key in new Outlook compose mode (see docs/outlook-quirks.md).
    item.notificationMessages.removeAsync(ENCRYPTION_STATUS_NOTIFICATION_KEY, () => {
      item.notificationMessages.replaceAsync(ENCRYPTION_STATUS_NOTIFICATION_KEY, details, () => {
        log(`banner set: encryptOn=${encryptOn}`);
        event.completed();
      });
    });
  };

  item.internetHeaders.getAsync([HEADER_ENCRYPT_ON_SEND], (hdrRes) => {
    let desired: boolean;
    let needsSeed: boolean;
    if (hdrRes.status === Office.AsyncResultStatus.Succeeded) {
      const v = hdrRes.value[HEADER_ENCRYPT_ON_SEND];
      if (v === "true") {
        desired = true;
        needsSeed = false;
      } else if (v === "false") {
        desired = false;
        needsSeed = false;
      } else {
        desired = globalDefault;
        needsSeed = true;
      }
    } else {
      // Header read failed. We have no way to know whether a prior
      // value already exists, so we still attempt to seed: if the write
      // succeeds the banner and OnMessageSend agree on the new value;
      // if it fails, we MUST paint "off" — anything else would be
      // lying to the user, because OnMessageSend reads only the header
      // and an unwritten/unreadable header releases the send in
      // cleartext.
      desired = globalDefault;
      needsSeed = true;
    }
    log(
      `desired=${desired} needsSeed=${needsSeed} ` +
        `header=${hdrRes.value?.[HEADER_ENCRYPT_ON_SEND] ?? "<absent>"} ` +
        `getStatus=${hdrRes.status}`
    );

    if (!needsSeed) {
      applyBanner(desired);
      return;
    }
    // Seed the header so OnMessageSend has a definite per-draft answer
    // even if the user never opens the taskpane. If the write fails,
    // paint "off" so the banner cannot lie about what will happen on
    // Send.
    item.internetHeaders.setAsync(
      { [HEADER_ENCRYPT_ON_SEND]: desired ? "true" : "false" },
      (setRes) => {
        const effective = setRes.status === Office.AsyncResultStatus.Succeeded ? desired : false;
        if (effective !== desired) {
          log(`header seed failed (status=${setRes.status}); painting banner as off`);
        }
        applyBanner(effective);
      }
    );
  });
}

log("script loaded");
Office.onReady((info) => {
  log(`Office.onReady fired; host=${info?.host} platform=${info?.platform}`);
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  log("handlers associated");
});
