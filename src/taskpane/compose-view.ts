// Compose-mode taskpane view: encryption toggle, policy editor entry points,
// and the "Encrypt & Send" action that runs the SDK + Yivi flow inline.

import { PostGuard, buildMime, UploadSessionExpiredError } from "@e4a/pg-js";
import {
  getRecipients,
  getSubject,
  setSubject,
  getBody,
  setBody,
  getAttachmentsCompose,
  readComposeAttachmentBytes,
  removeAttachment,
  addBase64Attachment,
  saveItem,
  setItemHeaders,
  removeItemHeaders,
  getItemHeaders,
  getComposeFromAsync,
  showNotification,
  removeNotification,
} from "../lib/office-helpers";
import { toBase64 } from "../lib/encoding";
import { EMAIL_ATTRIBUTE_TYPE } from "../lib/attributes";
import { Policy, MimeAttachment } from "../lib/types";
import {
  PKG_URL,
  CRYPTIFY_URL,
  POSTGUARD_WEBSITE_URL,
  ADDIN_VERSION,
  clientHeaders,
} from "../lib/pkg-client";
import { POSTGUARD_ENCRYPTED_FILENAME } from "../lib/mime";
import { buildSignAttributes, getEncryptionEnabled } from "../lib/settings";
import { t } from "../lib/i18n";
import { stringifyError } from "../lib/stringify-error";
import {
  recordPendingUpload,
  clearPendingUpload,
  probeAndClearPendingUpload,
} from "../lib/pending-upload";
import { mountPolicyPanel } from "./policy-editor";
import { showView, setStatus, showError } from "./taskpane";

// Internet-header keys shared with the OnMessageSend handler. Custom header
// names must be x-prefixed.
//
// Per-draft "encrypt this message on send" flag. This header is the
// single source of truth at send time — the OnMessageSend handler only
// runs the encryption flow when it reads "true" here, so the user's
// global Encryption setting only acts as the default that
// OnNewMessageCompose seeds onto each new draft.
const HEADER_ENCRYPT_ON_SEND = "x-pg-encrypt-on-send";
// Comma-joined sorted list of lowercase To+Cc emails captured at encrypt
// time. The handler compares this against the message's current recipients
// to refuse sending an encrypted blob to anyone who wasn't in the policy.
const HEADER_ENCRYPTED_RECIPIENTS = "x-pg-encrypted-recipients";
// PostGuard interop marker, written to outbound encrypted messages. The
// Thunderbird addon writes the same header (background.ts:485) and uses it
// as the OnMessageRead filter for the Outlook add-in. Detection on the
// receive side is still primarily attachment + body armor, but the header
// is a third independent signal that survives any HTML sanitation OWA
// applies during send.
const HEADER_POSTGUARD = "x-postguard";
const POSTGUARD_VERSION = "0.1.0";

async function persistEncryptOnSend(value: boolean): Promise<void> {
  try {
    // saveItem() before and after the header write: the first ensures
    // the draft has an itemId, the second flushes the header change to
    // the server so the OnMessageSend handler sees it. The header is
    // always set to an explicit "true" or "false" so the send-side
    // handler never has to fall back to anything else.
    await saveItem();
    await setItemHeaders({ [HEADER_ENCRYPT_ON_SEND]: value ? "true" : "false" });
    await saveItem();
    console.log(`[pg] persisted encryptOnSend=${value}`);
  } catch (e) {
    console.error(`[pg] failed to persist encryptOnSend:`, e);
  }
}

async function persistEncryptedRecipients(value: string | null): Promise<void> {
  try {
    await saveItem();
    if (value !== null) {
      await setItemHeaders({ [HEADER_ENCRYPTED_RECIPIENTS]: value });
    } else {
      await removeItemHeaders([HEADER_ENCRYPTED_RECIPIENTS]);
    }
    await saveItem();
  } catch (_e) {
    // Best-effort. The handler also re-derives the current recipient list
    // and compares; a missing or stale header just biases toward blocking.
  }
}

function recipientsKey(): string {
  return [...state.recipients.to, ...state.recipients.cc]
    .map((e) => e.toLowerCase().trim())
    .filter(Boolean)
    .sort()
    .join(",");
}

interface ComposeState {
  encrypt: boolean;
  policy: Policy;
  recipients: { to: string[]; cc: string[]; bcc: string[] };
  busy: boolean;
  // Set after a successful encrypt run; used to label the action button
  // "Re-encrypt" and disable it until something policy-relevant changes.
  encrypted: boolean;
  // Captured before encryption so a "Re-encrypt" can restore the draft
  // body and remove the previous encrypted attachment, then re-run from
  // scratch instead of double-encrypting the envelope.
  preEncryptBody: string | null;
  encryptedAttachmentId: string | null;
  // Hash of the policy-relevant inputs at last successful encryption.
  // Compared against `relevantStateString()` to decide whether the user
  // changed something since.
  encryptedSnapshot: string | null;
  // Last value we wrote to the x-pg-encrypted-recipients header so we only
  // re-write when it actually needs to change (renderToggleUI runs many
  // times between events). null means the header is currently cleared.
  encryptedRecipientsHeader: string | null;
}

const state: ComposeState = {
  encrypt: false,
  policy: {},
  recipients: { to: [], cc: [], bcc: [] },
  busy: false,
  encrypted: false,
  preEncryptBody: null,
  encryptedAttachmentId: null,
  encryptedSnapshot: null,
  encryptedRecipientsHeader: null,
};

// Single notification key for the encryption-status banner.
const ENCRYPTION_STATUS_NOTIFICATION_KEY = "postguard-encryption-status";

// Show the persistent in-message banner that mirrors the toggle. The
// taskpane is not always open while the user composes, so this banner is
// the user-visible "PostGuard is on/off" indicator on the message itself.
//
// Implementation note: new Outlook's notificationMessages.replaceAsync
// silently no-ops in compose mode when called with the same key — the
// callback reports success but the visible banner text doesn't change.
// Remove the existing entry first and then add the new one so the
// renderer actually re-paints. Best-effort throughout; a
// notificationMessages failure shouldn't break compose.
async function syncEncryptionBanner(): Promise<void> {
  try {
    const message = state.encrypt
      ? t("composeEncryptionOnBanner")
      : t("composeEncryptionOffBanner");
    await removeNotification(ENCRYPTION_STATUS_NOTIFICATION_KEY);
    await showNotification(ENCRYPTION_STATUS_NOTIFICATION_KEY, message, { persistent: true });
  } catch (_e) {
    // ignore
  }
}

// Stringified form of everything that affects the encrypted output. If this
// changes after a successful encrypt, the message no longer matches the
// current intent and Re-encrypt should be enabled. Sender attributes are
// configured in Settings (roaming) and don't participate here — changing a
// prefill doesn't invalidate the already-sealed message.
function relevantStateString(): string {
  return JSON.stringify({
    to: [...state.recipients.to].sort(),
    cc: [...state.recipients.cc].sort(),
    policy: state.policy,
  });
}

export async function mountComposeView(): Promise<void> {
  showView("compose");

  // Probe any leftover recovery token from a prior interrupted send.
  // Diagnostic only — pg-js 1.8.0 doesn't yet accept a pre-resumed
  // FileState back into createEnvelope, so the probe just drops stale
  // entries. See issue #82.
  void probeAndClearPendingUpload("roaming", CRYPTIFY_URL).catch(() => undefined);

  const toggle = byId<HTMLInputElement>("pg-toggle-encrypt");
  const bccWarning = byId<HTMLElement>("pg-bcc-warning");
  const manageTitle = byId<HTMLElement>("pg-manage-title");
  const btnEncryptSend = byId<HTMLButtonElement>("pg-btn-encrypt-send");

  manageTitle.textContent = t("manageAccess");
  btnEncryptSend.textContent = t("encryptAndSend");

  // The Encrypt & Send button is the Mac-only workaround for the
  // OnMessageSend launchevent dialog being broken on Outlook for Mac
  // (office-js#6677). Every other client opens the dialog directly
  // when the user hits Outlook's native Send, so the button is just
  // confusing UX there. Hide unless platform is Mac.
  btnEncryptSend.hidden = Office.context.platform !== Office.PlatformType.Mac;

  toggle.addEventListener("change", () => {
    state.encrypt = toggle.checked;
    // Write the per-draft header. The OnMessageSend handler reads only
    // this header — never the global setting — so a PostGuard outage
    // can never block an unencrypted send. The global Encryption
    // setting in the Settings view changes only the default for new
    // drafts (seeded by OnNewMessageCompose).
    void persistEncryptOnSend(state.encrypt);
    void syncEncryptionBanner();
    renderToggleUI();
    renderPolicyPanels();
  });

  btnEncryptSend.addEventListener("click", () => {
    if (state.busy) return;
    void encryptAndPrepareSend();
  });

  // Escape hatch out of the Yivi view. yivi-web shows a "cancelled" red X
  // inline when the user declines in the app and pg-js's promise behavior
  // around cancellation isn't fully reliable, so the user can stall here
  // without ever seeing our error view. This Cancel button always works.
  const btnYiviCancel = byId<HTMLButtonElement>("pg-btn-yivi-cancel");
  btnYiviCancel.textContent = t("policyEditorCancel");
  btnYiviCancel.addEventListener("click", () => {
    document.getElementById("yivi-web-form")!.innerHTML = "";
    state.busy = false;
    setStatus("");
    showView("compose");
  });

  // The per-draft header is authoritative. OnNewMessageCompose seeds
  // it from the mailbox-wide default on compose open, but the user may
  // have opened the taskpane before OnNewMessageCompose ran (or on a
  // draft that predates it), so seed here too if missing.
  try {
    const headers = await getItemHeaders([HEADER_ENCRYPT_ON_SEND]);
    const v = headers[HEADER_ENCRYPT_ON_SEND];
    if (v === "true") {
      state.encrypt = true;
    } else if (v === "false") {
      state.encrypt = false;
    } else {
      state.encrypt = getEncryptionEnabled();
      void persistEncryptOnSend(state.encrypt);
    }
  } catch (_e) {
    // Header read failed — fall back to the mailbox-wide default; the
    // user can still flip the toggle which will write the header.
    state.encrypt = getEncryptionEnabled();
  }

  await refreshRecipients();
  renderToggleUI();
  renderPolicyPanels();
  void syncEncryptionBanner();
  bccWarning.hidden = state.recipients.bcc.length === 0 || !state.encrypt;

  // Live recipient updates (Mailbox 1.7+). Without this the toggle UI is
  // stuck in whatever state the recipient lists were in at mount time.
  const item = Office.context.mailbox.item as Office.MessageCompose;
  item.addHandlerAsync(Office.EventType.RecipientsChanged, () => {
    void (async () => {
      await refreshRecipients();
      renderToggleUI();
      // Re-mount the manage panel so newly added/removed recipients show up
      // (or disappear) without needing a taskpane reopen.
      renderPolicyPanels();
    })();
  });
}

function renderPolicyPanels(): void {
  const manageSection = byId<HTMLElement>("pg-manage-section");
  const manageContainer = byId<HTMLElement>("pg-manage-panel");

  // When encryption is off the recipient policy doesn't apply — collapse
  // the section so the compose view stays uncluttered. Sender attributes
  // are configured in Settings now, not here.
  if (!state.encrypt) {
    manageSection.hidden = true;
    return;
  }
  manageSection.hidden = false;

  const recipients = [...state.recipients.to, ...state.recipients.cc];
  if (recipients.length === 0) {
    manageContainer.innerHTML = `<p class="pg-subtitle">${t("composeNoRecipients")}</p>`;
    return;
  }
  mountPolicyPanel(manageContainer, {
    emails: recipients,
    initialPolicy: state.policy,
    onChange: (next) => {
      state.policy = next;
      // Ensure email is always populated even if the user managed to clear it.
      for (const [email, attrs] of Object.entries(state.policy)) {
        if (!attrs.some((a) => a.t === EMAIL_ATTRIBUTE_TYPE)) {
          attrs.unshift({ t: EMAIL_ATTRIBUTE_TYPE, v: email });
        }
      }
      // Re-render toggle/button so the encrypt button reflects validation
      // (e.g. empty required-attribute values block encryption per #57).
      renderToggleUI();
    },
  });
}

function hasMissingPolicyValues(policy: Policy): boolean {
  for (const attrs of Object.values(policy)) {
    for (const a of attrs) {
      if (a.t === EMAIL_ATTRIBUTE_TYPE) continue;
      if (a.optional) continue;
      if (!a.v || a.v.trim().length === 0) return true;
    }
  }
  return false;
}

function renderToggleUI(): void {
  const toggle = byId<HTMLInputElement>("pg-toggle-encrypt");
  const toggleLabel = byId<HTMLElement>("pg-toggle-label");
  const btnEncryptSend = byId<HTMLButtonElement>("pg-btn-encrypt-send");
  const bccWarning = byId<HTMLElement>("pg-bcc-warning");

  toggle.checked = state.encrypt;
  toggleLabel.textContent = state.encrypt
    ? t("composeSwitchBarEnabled")
    : t("composeSwitchBarDisabled");

  const hasRecipients = state.recipients.to.length + state.recipients.cc.length > 0;
  const bccPresent = state.recipients.bcc.length > 0;
  const policyHasMissingValues = hasMissingPolicyValues(state.policy);

  // Re-encrypt mode: after a successful encryption, the button is only
  // useful if recipients/policy/sign attributes have drifted from what's
  // already on the draft. Otherwise re-clicking would just rebuild the
  // exact same envelope.
  const needsReencrypt = state.encrypted && relevantStateString() !== state.encryptedSnapshot;
  btnEncryptSend.textContent = state.encrypted ? t("reencryptAndSend") : t("encryptAndSend");
  btnEncryptSend.disabled =
    !state.encrypt ||
    !hasRecipients ||
    bccPresent ||
    policyHasMissingValues ||
    (state.encrypted && !needsReencrypt);
  if (state.encrypt && policyHasMissingValues) {
    setStatus(t("composePolicyValueMissing"), "error");
  }

  // Sync the x-pg-encrypted-recipients header to the current state. It
  // should hold the recipient list when the encryption is current, and be
  // cleared when state has drifted — so the OnMessageSend handler refuses
  // to send a now-stale ciphertext. Reverting a change re-stamps the
  // header, which re-allows sending without forcing a re-encrypt.
  if (state.encrypted) {
    const expected = needsReencrypt ? null : recipientsKey();
    if (state.encryptedRecipientsHeader !== expected) {
      state.encryptedRecipientsHeader = expected;
      void persistEncryptedRecipients(expected);
    }
  }

  if (bccPresent && state.encrypt) {
    bccWarning.hidden = false;
    bccWarning.textContent = t("composeBccWarning");
  } else {
    bccWarning.hidden = true;
  }
}

async function refreshRecipients(): Promise<void> {
  const [toR, ccR, bccR] = await Promise.all([
    getRecipients("to"),
    getRecipients("cc"),
    getRecipients("bcc"),
  ]);
  state.recipients.to = toR.map((r) => r.emailAddress.toLowerCase());
  state.recipients.cc = ccR.map((r) => r.emailAddress.toLowerCase());
  state.recipients.bcc = bccR.map((r) => r.emailAddress.toLowerCase());

  // Drop policy entries for emails no longer present.
  const all = new Set([...state.recipients.to, ...state.recipients.cc]);
  for (const k of Object.keys(state.policy)) {
    if (!all.has(k)) delete state.policy[k];
  }
  // Seed default (email-only) policy for new recipients.
  for (const email of all) {
    if (!state.policy[email]) {
      state.policy[email] = [{ t: EMAIL_ATTRIBUTE_TYPE, v: email }];
    }
  }
}

async function encryptAndPrepareSend(): Promise<void> {
  state.busy = true;
  setStatus(t("encrypting"));
  try {
    await refreshRecipients();
    if (state.recipients.bcc.length > 0) {
      throw new Error(t("composeBccWarning"));
    }
    if (state.recipients.to.length + state.recipients.cc.length === 0) {
      throw new Error(t("composeNoRecipients"));
    }
    if (hasMissingPolicyValues(state.policy)) {
      throw new Error(t("composePolicyValueMissing"));
    }

    const senderEmail = await getComposeFromAsync();
    if (!senderEmail) throw new Error(t("composeNoSenderEmail"));

    // If we're re-encrypting an already-encrypted draft, roll back first so
    // we encrypt the original plaintext body and attachments instead of
    // re-encrypting the previous envelope on top of itself.
    if (state.encrypted) {
      if (state.preEncryptBody !== null) {
        await setBody(state.preEncryptBody);
      }
      if (state.encryptedAttachmentId !== null) {
        try {
          await removeAttachment(state.encryptedAttachmentId);
        } catch (_e) {
          // Best-effort — user may have removed it manually.
        }
      }
      state.encrypted = false;
      state.preEncryptBody = null;
      state.encryptedAttachmentId = null;
      state.encryptedSnapshot = null;
      // Clear the on-message header too so the handler treats this as
      // unencrypted from this point until the new ciphertext is stamped.
      if (state.encryptedRecipientsHeader !== null) {
        state.encryptedRecipientsHeader = null;
        await persistEncryptedRecipients(null);
      }
    }

    const subject = await getSubject();
    const html = await getBody(Office.CoercionType.Html);
    const attachments = await collectComposeAttachments();

    const mime = (await buildMime({
      from: senderEmail,
      to: state.recipients.to,
      cc: state.recipients.cc,
      subject,
      htmlBody: html,
      date: new Date(),
      attachments: attachments.map((a) => ({
        name: a.name,
        type: a.type,
        data: a.data,
      })),
    } as never)) as Uint8Array;

    showView("yivi");
    const yiviTitle = byId<HTMLElement>("pg-yivi-title");
    const yiviSubtitle = byId<HTMLElement>("pg-yivi-subtitle");
    yiviTitle.textContent = t("displayMessageTitleSign");
    yiviSubtitle.textContent = t("displayMessageQrPrefix");
    // Reset the yivi host so the SDK can mount fresh.
    document.getElementById("yivi-web-form")!.innerHTML = "";

    const pg = new PostGuard({
      pkgUrl: PKG_URL,
      cryptifyUrl: CRYPTIFY_URL,
      headers: clientHeaders(ADDIN_VERSION),
    } as never);

    const recipients = buildPgRecipients(pg);

    // Sign attributes come from Settings (per-mailbox roaming setting),
    // not per-draft state. Prefilled values become { t, v } (mandatory);
    // blank prefills become { t, optional: true } so the user can decide
    // inside the Yivi app at session time. Fixes #49 / #56.
    const signAttributes = buildSignAttributes();

    const sealed = pg.encrypt({
      sign: pg.sign.yivi({
        element: "#yivi-web-form",
        senderEmail,
        includeSender: true,
        attributes: signAttributes,
      } as never),
      recipients,
      data: mime,
    } as never);

    // pg-js 1.2.0+: the Cryptify upload is silent by default, so we
    // can let it run for tier 2/3 — the recipient sees a download link
    // in the body but no duplicate mail from Cryptify.
    //
    // senderAttributes is a display-only hint for the envelope template.
    // With optional sign attributes we don't know the disclosed values
    // until after the Yivi session, so we leave it unset — the SDK falls
    // back to whatever the signed envelope itself carries.
    const envelope = await pg.email.createEnvelope({
      sealed,
      from: senderEmail,
      websiteUrl: POSTGUARD_WEBSITE_URL,
      onUploadInit: (info: { uuid: string; recoveryToken: string }) =>
        recordPendingUpload("roaming", info),
    } as never);
    clearPendingUpload("roaming");

    await setSubject(envelope.subject);
    await setBody(envelope.htmlBody);

    // Remove the now-redundant plaintext attachments. They are bundled
    // inside the encrypted envelope.
    for (const a of attachments) {
      try {
        await removeAttachment(a.id);
      } catch (_e) {
        // Continue best-effort.
      }
    }

    // Tier 1/2: include the encrypted bytes locally as postguard.encrypted.
    // Tier 3: pg-js gave us no attachment (too large) — recipients fetch
    // via the Cryptify link in the body.
    let attachmentId: string | null = null;
    if (envelope.attachment) {
      const attBytes = new Uint8Array(await envelope.attachment.arrayBuffer());
      const attBase64 = toBase64(attBytes);
      attachmentId = await addBase64Attachment(POSTGUARD_ENCRYPTED_FILENAME, attBase64);
    }

    // Force a server-side save before handing back to the user. Without this,
    // clicking Send can race the upload of the (potentially multi-MB) encrypted
    // body + attachment, which new Outlook surfaces as a "PostGuard timed out"
    // Smart Alerts dialog after ~15s.
    setStatus("Saving encrypted draft…");
    await saveItem();

    // Snapshot the encrypted state so renderToggleUI() can detect when the
    // user changes recipients/policy/sign attrs and re-enable Re-encrypt.
    state.encrypted = true;
    state.preEncryptBody = html;
    state.encryptedAttachmentId = attachmentId;
    state.encryptedSnapshot = relevantStateString();

    // Stamp the recipient set into a header so the OnMessageSend handler can
    // refuse to send if the user adds a new recipient afterwards (the new
    // recipient wouldn't be in the policy and couldn't decrypt). At the same
    // time write the cross-addon x-postguard interop marker.
    const stampedRecipients = recipientsKey();
    state.encryptedRecipientsHeader = stampedRecipients;
    await saveItem();
    await setItemHeaders({
      [HEADER_ENCRYPTED_RECIPIENTS]: stampedRecipients,
      [HEADER_POSTGUARD]: POSTGUARD_VERSION,
    });
    await saveItem();

    showView("compose");
    renderToggleUI();
    setStatus("Encrypted. Click Send to deliver the message.");
    await showNotification("postguard-encrypted", "PostGuard: message encrypted, click Send.", {
      persistent: true,
    });
  } catch (err) {
    // pg-js raises UploadSessionExpiredError when cryptify's structured
    // 404 says the upload session is gone (TTL expired, server restart,
    // unknown UUID, or wrong recovery_token). Show a clearer message
    // instead of the raw pg-js diagnostic. See issue #82.
    let msg: string;
    if (err instanceof UploadSessionExpiredError) {
      msg = t("uploadSessionExpiredError");
    } else {
      const detail = stringifyError(err);
      msg = detail || t("encryptionError");
    }
    setStatus(msg, "error");
    showView("compose");
    showError(msg);
  } finally {
    state.busy = false;
  }
}

function buildPgRecipients(pg: PostGuard): unknown[] {
  const all = [...state.recipients.to, ...state.recipients.cc];
  return all.map((email) => {
    const builder = (
      pg as never as { recipient: { email: (e: string) => RecipientBuilder } }
    ).recipient.email(email);
    const policy = state.policy[email];
    if (policy) {
      for (const attr of policy) {
        if (attr.t === EMAIL_ATTRIBUTE_TYPE) continue;
        // Manage Access entries always carry a value (the policy editor
        // filters out empty rows before storing them on state.policy),
        // but AttributeRequest.v is typed as optional now that sign-side
        // entries omit it — narrow defensively for the typechecker.
        if (!attr.v) continue;
        builder.extraAttribute(attr.t, attr.v.toLowerCase());
      }
    }
    return builder;
  });
}

interface RecipientBuilder {
  extraAttribute(t: string, v: string): RecipientBuilder;
}

async function collectComposeAttachments(): Promise<(MimeAttachment & { id: string })[]> {
  const list = await getAttachmentsCompose();
  const out: (MimeAttachment & { id: string })[] = [];
  for (const a of list) {
    // Skip cloud attachments — we cannot read their bytes via Office.js.
    if (a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud) continue;
    try {
      const data = await readComposeAttachmentBytes(a.id);
      out.push({
        id: a.id,
        name: a.name,
        type: guessContentType(a.name),
        data,
      });
    } catch (_e) {
      // Swallow individual attachment read failures.
    }
  }
  return out;
}

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
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
