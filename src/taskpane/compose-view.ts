// Compose-mode taskpane view: encryption toggle, policy editor entry points,
// and the "Encrypt & Send" action that runs the SDK + Yivi flow inline.

import { PostGuard, buildMime } from "@e4a/pg-js";
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
  getSenderEmail,
  showNotification,
} from "../lib/office-helpers";
import { toBase64 } from "../lib/encoding";
import { EMAIL_ATTRIBUTE_TYPE } from "../lib/attributes";
import { Policy, AttributeRequest, MimeAttachment } from "../lib/types";
import { PKG_URL, CRYPTIFY_URL, POSTGUARD_WEBSITE_URL, clientHeaders } from "../lib/pkg-client";
import { POSTGUARD_ENCRYPTED_FILENAME } from "../lib/mime";
import { t } from "../lib/i18n";
import { mountPolicyPanel } from "./policy-editor";
import { showView, setStatus, showError } from "./taskpane";

const ADDIN_VERSION = "0.1.0";

// Internet-header keys shared with the OnMessageSend handler. Custom header
// names must be x-prefixed.
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
    // saveItem() before and after the header write: the first ensures the
    // draft has an itemId, the second flushes the header change to the
    // server so the OnMessageSend handler sees it.
    //
    // Always write an explicit "true" or "false" — if we removed the
    // header for the off state, a draft the user explicitly toggled off
    // would reopen as default-on (since "absent" means "no choice yet").
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
  signAttributes: AttributeRequest[];
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
  encrypt: true,
  policy: {},
  signAttributes: [],
  recipients: { to: [], cc: [], bcc: [] },
  busy: false,
  encrypted: false,
  preEncryptBody: null,
  encryptedAttachmentId: null,
  encryptedSnapshot: null,
  encryptedRecipientsHeader: null,
};

// Stringified form of everything that affects the encrypted output. If this
// changes after a successful encrypt, the message no longer matches the
// current intent and Re-encrypt should be enabled.
function relevantStateString(): string {
  return JSON.stringify({
    to: [...state.recipients.to].sort(),
    cc: [...state.recipients.cc].sort(),
    policy: state.policy,
    sign: state.signAttributes,
  });
}

export async function mountComposeView(): Promise<void> {
  showView("compose");

  const toggle = byId<HTMLInputElement>("pg-toggle-encrypt");
  const bccWarning = byId<HTMLElement>("pg-bcc-warning");
  const manageTitle = byId<HTMLElement>("pg-manage-title");
  const signTitle = byId<HTMLElement>("pg-sign-title");
  const btnEncryptSend = byId<HTMLButtonElement>("pg-btn-encrypt-send");

  manageTitle.textContent = t("manageAccess");
  signTitle.textContent = t("sign");
  btnEncryptSend.textContent = t("encryptAndSend");

  // The Encrypt & Send button is the Mac-only workaround for the
  // OnMessageSend launchevent dialog being broken on Outlook for Mac
  // (office-js#6677). Every other client opens the dialog directly
  // when the user hits Outlook's native Send, so the button is just
  // confusing UX there. Hide unless platform is Mac.
  btnEncryptSend.hidden = Office.context.platform !== Office.PlatformType.Mac;

  toggle.addEventListener("change", () => {
    state.encrypt = toggle.checked;
    void persistEncryptOnSend(state.encrypt);
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

  // Restore the toggle state from the per-draft header so a soft-block
  // round trip or a taskpane reopen doesn't lose the user's choice.
  // The header has three states:
  //   "true"  → user explicitly enabled
  //   "false" → user explicitly disabled
  //   absent  → never interacted; fall back to the default-on behaviour
  //             and persist "true" so the OnMessageSend handler sees the
  //             same intent the toggle visually shows.
  try {
    const headers = await getItemHeaders([HEADER_ENCRYPT_ON_SEND]);
    const v = headers[HEADER_ENCRYPT_ON_SEND];
    if (v === "true") {
      state.encrypt = true;
    } else if (v === "false") {
      state.encrypt = false;
    } else {
      state.encrypt = true;
      void persistEncryptOnSend(true);
    }
  } catch (_e) {
    // Header read failed — leave the default-on state alone.
  }

  await refreshRecipients();
  renderToggleUI();
  renderPolicyPanels();
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
  const signSection = byId<HTMLElement>("pg-sign-section");
  const manageContainer = byId<HTMLElement>("pg-manage-panel");
  const signContainer = byId<HTMLElement>("pg-sign-panel");

  // When encryption is off, the policies don't apply — collapse both
  // sections so the compose view stays uncluttered.
  if (!state.encrypt) {
    manageSection.hidden = true;
    signSection.hidden = true;
    return;
  }
  manageSection.hidden = false;
  signSection.hidden = false;

  const recipients = [...state.recipients.to, ...state.recipients.cc];
  if (recipients.length === 0) {
    manageContainer.innerHTML = `<p class="pg-subtitle">${t("composeNoRecipients")}</p>`;
  } else {
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
      },
    });
  }

  const senderEmail = getSenderEmail();
  if (!senderEmail) {
    signContainer.innerHTML = "";
    return;
  }
  // Sign editor is conceptually a single-recipient policy where the
  // "recipient" is the sender's own address.
  const signInitial: Policy = {
    [senderEmail]: [{ t: EMAIL_ATTRIBUTE_TYPE, v: senderEmail }, ...state.signAttributes],
  };
  mountPolicyPanel(signContainer, {
    emails: [senderEmail],
    initialPolicy: signInitial,
    onChange: (next) => {
      // signAttributes stores ONLY extras. pg.sign.yivi already takes
      // senderEmail as a top-level field; including email here as well
      // triggers a second email disclosure on the Yivi QR.
      state.signAttributes = (next[senderEmail] ?? []).filter((a) => a.t !== EMAIL_ATTRIBUTE_TYPE);
    },
  });
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

  // Re-encrypt mode: after a successful encryption, the button is only
  // useful if recipients/policy/sign attributes have drifted from what's
  // already on the draft. Otherwise re-clicking would just rebuild the
  // exact same envelope.
  const needsReencrypt = state.encrypted && relevantStateString() !== state.encryptedSnapshot;
  btnEncryptSend.textContent = state.encrypted ? t("reencryptAndSend") : t("encryptAndSend");
  btnEncryptSend.disabled =
    !state.encrypt || !hasRecipients || bccPresent || (state.encrypted && !needsReencrypt);

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

    const senderEmail = getSenderEmail();
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

    const sealed = pg.encrypt({
      sign: pg.sign.yivi({
        element: "#yivi-web-form",
        senderEmail,
        attributes: state.signAttributes.length ? state.signAttributes : undefined,
      } as never),
      recipients,
      data: mime,
    } as never);

    // pg-js 1.2.0+: the Cryptify upload is silent by default, so we
    // can let it run for tier 2/3 — the recipient sees a download link
    // in the body but no duplicate mail from Cryptify.
    const envelope = await pg.email.createEnvelope({
      sealed,
      from: senderEmail,
      websiteUrl: POSTGUARD_WEBSITE_URL,
      senderAttributes: state.signAttributes.map((a) => a.v),
    } as never);

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
    const msg = err instanceof Error ? err.message : t("encryptionError");
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
        if (attr.t !== EMAIL_ATTRIBUTE_TYPE) {
          builder.extraAttribute(attr.t, attr.v.toLowerCase());
        }
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
