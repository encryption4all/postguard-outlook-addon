/* global Office */

import "./taskpane.css";
import {
  PG_ATTACHMENT_NAME,
  PKG_URL,
  PG_CLIENT_HEADER,
  EMAIL_ATTRIBUTE_TYPE,
  toEmail,
  typeToImage,
  retrievePublicKey,
  retrieveVerificationKey,
  getUSK,
  checkJwtCache,
  storeJwtCache,
  secondsTill4AM,
  cleanUpCache,
  extractArmoredPayload,
} from "../utils";
import type { AttributeCon, Badge } from "../types";

// Module-level state
let pgWasm: typeof import("@e4a/pg-wasm") | null = null;
let masterPublicKey: string | null = null;
let masterVerificationKey: string | null = null;

// UI Helpers
function showSection(id: string): void {
  const sections = document.querySelectorAll(".pg-section");
  sections.forEach((s) => ((s as HTMLElement).style.display = "none"));
  const section = document.getElementById(id);
  if (section) section.style.display = "block";
}

function showError(message: string): void {
  const el = document.getElementById("pg-error-message");
  if (el) el.textContent = message;
  showSection("pg-error");
}

function renderBadges(containerId: string, badges: Badge[]): void {
  const container = document.getElementById(containerId);
  if (!container) return;
  container.innerHTML = "";

  const iconMap: Record<string, string> = {
    envelope: "Mail",
    phone: "Phone",
    personal: "Contact",
    education: "Education",
    health: "Health",
    calendar: "Calendar",
  };

  for (const badge of badges) {
    const el = document.createElement("span");
    el.className = "pg-badge";
    const iconName = iconMap[badge.type] || "Contact";
    el.innerHTML = `<i class="ms-Icon ms-Icon--${iconName} pg-badge-icon"></i>${badge.value}`;
    container.appendChild(el);
  }
}

// Initialize WASM module and keys (lazy, only when needed)
async function initPostGuard(): Promise<void> {
  if (pgWasm && masterPublicKey && masterVerificationKey) return;
  console.log("[PostGuard] Initializing...");
  const [pk, vk, mod] = await Promise.all([
    retrievePublicKey(),
    retrieveVerificationKey(),
    import("@e4a/pg-wasm"),
  ]);
  // Initialize the WASM module (default export is the init function)
  await mod.default();
  masterPublicKey = pk;
  masterVerificationKey = vk;
  pgWasm = mod;
  console.log("[PostGuard] Initialization complete.");
}

// Check if the current message has a postguard.encrypted attachment or armored body
async function detectEncryption(): Promise<{
  isEncrypted: boolean;
  attachmentId?: string;
  armoredBase64?: string;
}> {
  const item = Office.context.mailbox.item;
  if (!item) return { isEncrypted: false };

  // Check attachments (read-mode uses the .attachments property, not getAttachmentsAsync)
  let attachmentResult: { attachmentId?: string } = {};
  try {
    const attachments: Office.AttachmentDetails[] = (item as any).attachments ?? [];
    const pgAttachment = attachments.find((att) => att.name === PG_ATTACHMENT_NAME);
    if (pgAttachment) {
      attachmentResult = { attachmentId: pgAttachment.id };
    }
  } catch (e) {
    console.log("[PostGuard] detectEncryption attachment error:", e);
  }

  if (attachmentResult.attachmentId) {
    return { isEncrypted: true, attachmentId: attachmentResult.attachmentId };
  }

  // Fallback: check body for armored payload
  try {
    const bodyHtml = await new Promise<string>((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Html, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else reject(new Error("Failed to get body"));
      });
    });
    const armoredBase64 = extractArmoredPayload(bodyHtml);
    if (armoredBase64) {
      return { isEncrypted: true, armoredBase64 };
    }
  } catch (e) {
    console.log("[PostGuard] detectEncryption body check error:", e);
  }

  return { isEncrypted: false };
}

// Get the encrypted attachment content as base64
async function getAttachmentContent(attachmentId: string): Promise<string> {
  const item = Office.context.mailbox.item;
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

// Active Yivi session (so we can abort it)
let activeYiviSession: { abort: () => void } | null = null;

// Start inline Yivi authentication in the taskpane
async function startYiviAuth(con: AttributeCon, senderId?: string): Promise<string> {
  showSection("pg-yivi");

  // Show sender info if available
  if (senderId) {
    const senderInfo = document.getElementById("pg-sender-info");
    const senderEmail = document.getElementById("pg-sender-email");
    if (senderInfo) senderInfo.style.display = "block";
    if (senderEmail) senderEmail.textContent = senderId;
  }

  // Clear any previous Yivi form content
  const formEl = document.getElementById("yivi-web-form");
  if (formEl) formEl.innerHTML = "";

  // Shim process for Node.js polyfills used by yivi-client's dependencies
  if (typeof (window as any).process === "undefined") {
    (window as any).process = { env: {}, version: "", browser: true };
  }

  // Dynamically import Yivi modules
  const [YiviCore, YiviClient, YiviWeb] = await Promise.all([
    import("@privacybydesign/yivi-core"),
    import("@privacybydesign/yivi-client"),
    import("@privacybydesign/yivi-web"),
  ]);

  return new Promise<string>((resolve, reject) => {
    const yivi = new YiviCore.default({
      debugging: false,
      element: "#yivi-web-form",
      language: navigator.language.startsWith("nl") ? "nl" : "en",
      translations: {
        header: "",
        helper: "Scan with your Yivi app to decrypt",
      },
      state: {
        serverSentEvents: false,
        polling: {
          endpoint: "status",
          interval: 500,
          startState: "INITIALIZED",
        },
      },
      session: {
        url: PKG_URL,
        start: {
          url: (o: { url: string }) => `${o.url}/v2/request/start`,
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            ...PG_CLIENT_HEADER,
          },
          body: JSON.stringify({ con, validity: secondsTill4AM() }),
        },
        result: {
          url: (o: { url: string }, { sessionToken }: { sessionToken: string }) =>
            `${o.url}/v2/request/jwt/${sessionToken}`,
          headers: PG_CLIENT_HEADER,
          parseResponse: (r: Response) => r.text(),
        },
      },
    });

    yivi.use(YiviClient.default);
    yivi.use(YiviWeb.default);

    activeYiviSession = {
      abort: () => {
        try { yivi.abort(); } catch { /* ignore */ }
        reject(new Error("Yivi session cancelled"));
      },
    };

    // Cancel button
    const btnCancel = document.getElementById("btn-cancel-yivi");
    if (btnCancel) {
      btnCancel.onclick = () => {
        activeYiviSession?.abort();
        activeYiviSession = null;
        showSection("pg-encrypted");
      };
    }

    yivi
      .start()
      .then((jwt: string) => {
        activeYiviSession = null;
        resolve(jwt);
      })
      .catch((e: Error) => {
        activeYiviSession = null;
        reject(e);
      });
  });
}

// Main decryption flow — accepts attachment ID or raw base64 string
async function decryptMessage(source: { attachmentId: string } | { base64: string }): Promise<void> {
  showSection("pg-decrypting");

  try {
    await initPostGuard();

    let base64Content: string;
    if ("attachmentId" in source) {
      base64Content = await getAttachmentContent(source.attachmentId);
    } else {
      base64Content = source.base64;
    }
    const binaryString = atob(base64Content);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    const readable = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(bytes);
        controller.close();
      },
    });

    const unsealer = await pgWasm!.StreamUnsealer.new(readable, masterVerificationKey!);

    const userEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    const recipients = unsealer.inspect_header();
    const me = recipients.get(userEmail);

    if (!me) {
      throw Object.assign(new Error("Your email address was not found in the encryption recipients"), {
        name: "RecipientUnknownError",
      });
    }

    const keyRequest = Object.assign({}, me);
    let hints = me.con;

    hints = hints.map(({ t, v }: { t: string; v: string }) => {
      if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: userEmail };
      return { t, v };
    });

    keyRequest.con = keyRequest.con.map(({ t, v }: { t: string; v: string }) => {
      if (t === EMAIL_ATTRIBUTE_TYPE) return { t, v: userEmail };
      if (v === "" || v.includes("*")) return { t };
      return { t, v };
    });

    const senderEmail = Office.context.mailbox.item?.from?.emailAddress;
    const jwt = await checkJwtCache(hints).catch(() =>
      startYiviAuth(keyRequest.con, senderEmail)
    );

    const usk = await getUSK(jwt, keyRequest.ts);

    let decryptedData = "";
    const decoder = new TextDecoder();
    const writable = new WritableStream({
      write(chunk: Uint8Array) {
        decryptedData += decoder.decode(chunk, { stream: true });
      },
      close() {
        decryptedData += decoder.decode();
      },
    });

    const senderIdentity = await unsealer.unseal(userEmail, usk, writable);
    console.log("[PostGuard] Sender verification:", senderIdentity);

    await storeJwtCache(hints, jwt);
    displayDecryptedContent(decryptedData, senderIdentity);
  } catch (e: unknown) {
    console.error("[PostGuard] Decryption error:", e);
    if (e instanceof Error) {
      if (e.name === "RecipientUnknownError") {
        showError("Your email address is not among the recipients of this encrypted message.");
      } else if (e.name === "OperationError") {
        showError("Decryption failed. The message could not be decrypted with your credentials.");
      } else if (e.message === "Dialog was closed" || e.message === "Yivi session cancelled") {
        showSection("pg-encrypted");
      } else {
        showError(e.message);
      }
    } else {
      showError("An unexpected error occurred during decryption.");
    }
  }
}

function displayDecryptedContent(mimeData: string, senderIdentity: unknown): void {
  const subjectMatch = mimeData.match(/^Subject:\s*(.+)$/im);
  const subject = subjectMatch ? subjectMatch[1].trim() : "(no subject)";

  const headerEndIndex = mimeData.indexOf("\r\n\r\n");
  let body = headerEndIndex !== -1 ? mimeData.substring(headerEndIndex + 4) : mimeData;

  const contentTypeMatch = mimeData.match(/^Content-Type:\s*multipart\/mixed;\s*boundary="?([^"\r\n]+)"?/im);
  if (contentTypeMatch) {
    const boundary = contentTypeMatch[1];
    const parts = body.split(`--${boundary}`);
    if (parts.length > 1) {
      const firstPart = parts[1];
      const partHeaderEnd = firstPart.indexOf("\r\n\r\n");
      body = partHeaderEnd !== -1 ? firstPart.substring(partHeaderEnd + 4) : firstPart;
    }
  }

  const subjectEl = document.getElementById("pg-subject-text");
  if (subjectEl) subjectEl.textContent = subject;

  const bodyEl = document.getElementById("pg-body-content");
  if (bodyEl) {
    if (body.includes("<html") || body.includes("<HTML") || body.includes("<div") || body.includes("<p")) {
      bodyEl.innerHTML = body;
    } else {
      bodyEl.textContent = body;
    }
  }

  if (senderIdentity) {
    const identity = senderIdentity as { public: { con: Array<{ t: string; v: string }> }; private?: { con: Array<{ t: string; v: string }> } };
    const privBadges = identity.private?.con ?? [];
    const badges: Badge[] = [...identity.public.con, ...privBadges].map(({ t, v }) => ({
      type: typeToImage(t),
      value: v,
    }));
    renderBadges("pg-sender-badges", badges);
  }

  showSection("pg-decrypted");
}

// Main initialization
// Handle the current message: check cache from launch event, detect encryption, etc.
async function handleCurrentItem(): Promise<void> {
  // Don't block on PKG init - detect encryption first, init WASM only when user decrypts
  try {
    const { isEncrypted, attachmentId, armoredBase64 } = await detectEncryption();
    console.log("[PostGuard] Encryption detected:", isEncrypted);

    if (isEncrypted && (attachmentId || armoredBase64)) {
      showSection("pg-encrypted");

      const source = attachmentId ? { attachmentId } : { base64: armoredBase64! };
      const btnDecrypt = document.getElementById("btn-decrypt");
      if (btnDecrypt) {
        btnDecrypt.onclick = () => decryptMessage(source);
      }
    } else {
      // Check X-PostGuard header
      const item = Office.context.mailbox.item;
      if (item && item.getAllInternetHeadersAsync) {
        item.getAllInternetHeadersAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const headers = result.value;
            if (headers.toLowerCase().includes("x-postguard")) {
              showSection("pg-was-encrypted");
              return;
            }
          }
          showSection("pg-not-encrypted");
        });
      } else {
        showSection("pg-not-encrypted");
      }
    }
  } catch (e) {
    console.error("[PostGuard] Init error:", e);
    showSection("pg-not-encrypted");
  }
}

// Main initialization
Office.onReady(async (info) => {
  console.log("[PostGuard] Office.onReady fired, host:", info.host);

  cleanUpCache();

  await handleCurrentItem();

  // Re-run when the user switches messages (pinned taskpane)
  Office.context.mailbox.addHandlerAsync(
    Office.EventType.ItemChanged,
    () => handleCurrentItem()
  );

  const btnRetry = document.getElementById("btn-retry");
  if (btnRetry) {
    btnRetry.onclick = async () => {
      const { isEncrypted, attachmentId, armoredBase64 } = await detectEncryption();
      if (isEncrypted) {
        const retrySource = attachmentId ? { attachmentId } : { base64: armoredBase64! };
        await decryptMessage(retrySource);
      }
    };
  }
});
