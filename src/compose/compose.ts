/* global Office */

import "./compose.css";
import {
  EMAIL_ATTRIBUTE_TYPE,
  toEmail,
  typeToImage,
  PKG_URL,
  PG_CLIENT_HEADER,
  secondsTill4AM,
  checkJwtCache,
  storeJwtCache,
} from "../utils";
import type { Policy, AttributeCon, Badge } from "../types";

// Compose state persisted in the task pane session
let encryptionEnabled = false;
let recipientPolicy: Policy = {};
let signIdentity: AttributeCon | null = null;

// Known Yivi attribute types for the attribute selector
const ATTRIBUTE_TYPES: { type: string; label: string; category: string }[] = [
  { type: "pbdf.sidn-pbdf.email.email", label: "Email address", category: "Contact" },
  { type: "pbdf.sidn-pbdf.mobilenumber.mobilenumber", label: "Mobile number", category: "Contact" },
  { type: "pbdf.pbdf.surfnet-2.id", label: "SURFnet ID (education)", category: "Education" },
  { type: "pbdf.nuts.agb.agbcode", label: "AGB code (healthcare)", category: "Healthcare" },
  { type: "pbdf.gemeente.personalData.dateofbirth", label: "Date of birth", category: "Personal" },
];

function showSection(id: string): void {
  const sections = document.querySelectorAll(".pg-section");
  sections.forEach((s) => ((s as HTMLElement).style.display = "none"));
  const section = document.getElementById(id);
  if (section) section.style.display = "block";
}

// Get recipients from compose item
async function getRecipients(): Promise<string[]> {
  const item = Office.context.mailbox.item;
  if (!item) return [];

  const toRecipients = await new Promise<Office.EmailAddressDetails[]>((resolve) => {
    item.to.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
    });
  });

  const ccRecipients = await new Promise<Office.EmailAddressDetails[]>((resolve) => {
    item.cc.getAsync((result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []);
    });
  });

  return [...toRecipients, ...ccRecipients].map((r) => toEmail(r.emailAddress));
}

// Check for BCC recipients
async function hasBccRecipients(): Promise<boolean> {
  const item = Office.context.mailbox.item;
  if (!item || !item.bcc) return false;

  return new Promise<boolean>((resolve) => {
    item.bcc.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value.length > 0);
      } else {
        resolve(false);
      }
    });
  });
}

// Render recipient list with policy indicators
function renderRecipients(recipients: string[]): void {
  const list = document.getElementById("pg-recipient-list");
  const noRecipients = document.getElementById("pg-no-recipients");
  const attrCard = document.getElementById("pg-attribute-card");

  if (!list) return;
  list.innerHTML = "";

  if (recipients.length === 0) {
    if (noRecipients) noRecipients.style.display = "block";
    if (attrCard) attrCard.style.display = "none";
    return;
  }

  if (noRecipients) noRecipients.style.display = "none";
  if (attrCard) attrCard.style.display = "block";

  for (const email of recipients) {
    const li = document.createElement("li");
    li.className = "pg-recipient-item";

    const emailSpan = document.createElement("span");
    emailSpan.className = "pg-recipient-email";
    emailSpan.innerHTML = `<i class="ms-Icon ms-Icon--Mail"></i> ${email}`;

    const policy = recipientPolicy[email];
    const policySpan = document.createElement("span");
    policySpan.className = "pg-recipient-policy";
    if (policy && policy.length > 0) {
      policySpan.textContent = policy.map((a) => a.t.split(".").pop()).join(", ");
    } else {
      policySpan.textContent = "email";
    }

    li.appendChild(emailSpan);
    li.appendChild(policySpan);
    list.appendChild(li);

    // Initialize default policy (email) if not set
    if (!recipientPolicy[email]) {
      recipientPolicy[email] = [];
    }
  }

  renderAttributeConfig(recipients);
}

// Render attribute configuration for recipients
function renderAttributeConfig(recipients: string[]): void {
  const container = document.getElementById("pg-attribute-config");
  if (!container) return;
  container.innerHTML = "";

  for (const email of recipients) {
    const group = document.createElement("div");
    group.className = "pg-attr-group";

    const header = document.createElement("div");
    header.className = "pg-attr-group-header ms-font-s";
    header.textContent = email;
    group.appendChild(header);

    // Default option: email-based
    const defaultOption = createAttributeOption(email, EMAIL_ATTRIBUTE_TYPE, "Email address (default)", true);
    group.appendChild(defaultOption);

    // Additional attribute options
    for (const attr of ATTRIBUTE_TYPES) {
      if (attr.type === EMAIL_ATTRIBUTE_TYPE) continue;
      const option = createAttributeOption(email, attr.type, attr.label, false);
      group.appendChild(option);
    }

    container.appendChild(group);
  }
}

function createAttributeOption(email: string, attrType: string, label: string, isDefault: boolean): HTMLElement {
  const div = document.createElement("div");
  div.className = "pg-attr-option";

  const radio = document.createElement("input");
  radio.type = "radio";
  radio.name = `attr-${email}`;
  radio.value = attrType;
  radio.id = `attr-${email}-${attrType}`;

  // Check if this is the currently selected policy
  const currentPolicy = recipientPolicy[email];
  if (isDefault && (!currentPolicy || currentPolicy.length === 0)) {
    radio.checked = true;
  } else if (currentPolicy?.some((a) => a.t === attrType)) {
    radio.checked = true;
  }

  radio.addEventListener("change", () => {
    if (attrType === EMAIL_ATTRIBUTE_TYPE) {
      // Default: just use email
      recipientPolicy[email] = [];
    } else {
      recipientPolicy[email] = [{ t: attrType }];
    }
    // Save state
    saveComposeState();
  });

  const labelEl = document.createElement("label");
  labelEl.htmlFor = radio.id;
  labelEl.textContent = label;

  div.appendChild(radio);
  div.appendChild(labelEl);
  return div;
}

// Active Yivi session (so we can abort it)
let activeYiviSession: { abort: () => void } | null = null;

// Start inline Yivi authentication for signing identity
async function startSignYiviAuth(): Promise<void> {
  const userEmail = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
  const con: AttributeCon = [{ t: EMAIL_ATTRIBUTE_TYPE, v: userEmail }];

  showSection("pg-yivi");

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

  return new Promise<void>((resolve, reject) => {
    const yivi = new YiviCore.default({
      debugging: false,
      element: "#yivi-web-form",
      language: navigator.language.startsWith("nl") ? "nl" : "en",
      translations: {
        header: "",
        helper: "Scan with your Yivi app to sign",
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
        showSection("pg-compose-main");
      };
    }

    yivi
      .start()
      .then(async (jwt: string) => {
        activeYiviSession = null;
        signIdentity = con;
        await storeJwtCache(con, jwt);
        renderSignBadges();
        showSection("pg-compose-main");
        resolve();
      })
      .catch((e: Error) => {
        activeYiviSession = null;
        reject(e);
      });
  });
}

function renderSignBadges(): void {
  const container = document.getElementById("pg-sign-badges");
  if (!container || !signIdentity) return;
  container.innerHTML = "";

  const iconMap: Record<string, string> = {
    envelope: "Mail",
    phone: "Phone",
    personal: "Contact",
    education: "Education",
    health: "Health",
    calendar: "Calendar",
  };

  for (const attr of signIdentity) {
    const badge: Badge = { type: typeToImage(attr.t), value: attr.v || attr.t.split(".").pop() || "" };
    const el = document.createElement("span");
    el.className = "pg-badge";
    const iconName = iconMap[badge.type] || "Contact";
    el.innerHTML = `<i class="ms-Icon ms-Icon--${iconName}"></i> ${badge.value}`;
    container.appendChild(el);
  }
}

// Save compose state to sessionStorage so the onSend handler can access it
function saveComposeState(): void {
  const state = {
    encrypt: encryptionEnabled,
    policy: recipientPolicy,
    signId: signIdentity,
  };
  sessionStorage.setItem("pg-compose-state", JSON.stringify(state));
}

// Load compose state
function loadComposeState(): void {
  const saved = sessionStorage.getItem("pg-compose-state");
  if (saved) {
    try {
      const state = JSON.parse(saved);
      encryptionEnabled = state.encrypt || false;
      recipientPolicy = state.policy || {};
      signIdentity = state.signId || null;
    } catch {
      // ignore
    }
  }
}

// Update UI based on encryption toggle
function updateEncryptionUI(): void {
  const panel = document.getElementById("pg-encrypt-panel");
  const noPanel = document.getElementById("pg-no-encrypt-panel");
  const lockIcon = document.getElementById("pg-lock-icon");
  const toggleText = document.getElementById("pg-toggle-text");

  if (panel) panel.style.display = encryptionEnabled ? "block" : "none";
  if (noPanel) noPanel.style.display = encryptionEnabled ? "none" : "block";
  if (lockIcon) {
    lockIcon.className = encryptionEnabled
      ? "ms-Icon ms-Icon--LockSolid ms-font-l"
      : "ms-Icon ms-Icon--Lock ms-font-l";
    lockIcon.style.color = encryptionEnabled ? "#006ef4" : "";
  }
  if (toggleText) {
    toggleText.style.fontWeight = encryptionEnabled ? "600" : "400";
  }

  // Set notification on the mail item
  const item = Office.context.mailbox.item;
  if (item && item.notificationMessages) {
    if (encryptionEnabled) {
      item.notificationMessages.replaceAsync("pgEncrypt", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "PostGuard encryption is enabled for this email.",
        icon: "Icon.16x16",
        persistent: false,
      });
    } else {
      item.notificationMessages.removeAsync("pgEncrypt");
    }
  }
}

// Main init
Office.onReady(async (info) => {
  console.log("[PostGuard Compose] Office.onReady fired, host:", info.host);

  loadComposeState();
  showSection("pg-compose-main");

  // Set toggle state
  const toggle = document.getElementById("pg-encrypt-toggle") as HTMLInputElement;
  if (toggle) {
    toggle.checked = encryptionEnabled;
    toggle.addEventListener("change", async () => {
      encryptionEnabled = toggle.checked;
      saveComposeState();
      updateEncryptionUI();

      if (encryptionEnabled) {
        const recipients = await getRecipients();
        renderRecipients(recipients);

        // Check BCC
        const hasBcc = await hasBccRecipients();
        const bccWarning = document.getElementById("pg-bcc-warning");
        if (bccWarning) bccWarning.style.display = hasBcc ? "flex" : "none";
      }
    });
  }

  updateEncryptionUI();

  // If encryption was already enabled, load recipients
  if (encryptionEnabled) {
    const recipients = await getRecipients();
    renderRecipients(recipients);
    renderSignBadges();
  }

  // Refresh recipients button
  const btnRefresh = document.getElementById("btn-refresh-recipients");
  if (btnRefresh) {
    btnRefresh.addEventListener("click", async () => {
      const recipients = await getRecipients();
      renderRecipients(recipients);

      const hasBcc = await hasBccRecipients();
      const bccWarning = document.getElementById("pg-bcc-warning");
      if (bccWarning) bccWarning.style.display = hasBcc ? "flex" : "none";
    });
  }

  // Configure signing identity
  const btnSign = document.getElementById("btn-configure-sign");
  if (btnSign) {
    btnSign.addEventListener("click", async () => {
      try {
        await startSignYiviAuth();
      } catch (e) {
        console.log("[PostGuard] Sign auth error:", e);
        showSection("pg-compose-main");
      }
    });
  }
});
