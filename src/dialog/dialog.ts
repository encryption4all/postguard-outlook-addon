/* global Office */

import "./dialog.css";
import * as YiviCore from "@privacybydesign/yivi-core";
import * as YiviClient from "@privacybydesign/yivi-client";
import * as YiviWeb from "@privacybydesign/yivi-web";

interface DialogData {
  hostname: string;
  header: Record<string, string>;
  con: Array<{ t: string; v?: string }>;
  sort: string;
  hints?: Array<{ t: string; v?: string }>;
  senderId?: string;
  validity: number;
}

function getDialogData(): DialogData {
  const params = new URLSearchParams(window.location.search);
  const dataStr = params.get("data");
  if (!dataStr) throw new Error("No data parameter found");
  return JSON.parse(decodeURIComponent(dataStr));
}

function fillAttributeTable(data: DialogData): void {
  const table = document.querySelector("#pg-attribute-table tbody");
  if (!table) return;

  const hints = data.hints || data.con;
  for (const { t, v } of hints) {
    const row = document.createElement("tr");
    const tdType = document.createElement("td");
    const tdValue = document.createElement("td");

    // Use the last part of the attribute type as a readable name
    const readableName = t.split(".").pop() || t;
    tdType.textContent = readableName;
    tdValue.textContent = v || "";
    tdValue.className = "value";

    row.appendChild(tdType);
    row.appendChild(tdValue);
    table.appendChild(row);
  }

  const attrInfo = document.getElementById("pg-attributes-info");
  if (attrInfo && hints.length > 0) attrInfo.style.display = "block";
}

function initializeDialog(): void {
  let data: DialogData;
  try {
    data = getDialogData();
  } catch (e) {
    console.error("[PostGuard Dialog] Failed to get dialog data:", e);
    return;
  }

  // Set title based on sort
  const title = document.getElementById("pg-dialog-title");
  const heading = document.getElementById("pg-dialog-heading");

  if (data.sort === "Decryption") {
    if (title) title.textContent = "Decrypt Message";
    if (heading) heading.textContent = "Scan the QR code below with your Yivi app to prove your identity and decrypt this message.";
  } else {
    if (title) title.textContent = "Sign Identity";
    if (heading) heading.textContent = "Scan the QR code below with your Yivi app to attach your verified identity to this message.";
  }

  // Show sender info
  if (data.senderId) {
    const senderInfo = document.getElementById("pg-sender-info");
    const senderEmail = document.getElementById("pg-sender-email");
    if (senderInfo) senderInfo.style.display = "block";
    if (senderEmail) senderEmail.textContent = data.senderId;
  }

  // Fill attribute table
  fillAttributeTable(data);

  // Initialize Yivi
  const yivi = new YiviCore({
    debugging: false,
    element: "#yivi-web-form",
    language: navigator.language.startsWith("nl") ? "nl" : "en",
    translations: {
      header: "",
      helper: data.sort === "Decryption"
        ? "Scan with your Yivi app to decrypt"
        : "Scan with your Yivi app to sign",
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
      url: data.hostname,
      start: {
        url: (o: { url: string }) => `${o.url}/v2/request/start`,
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          ...data.header,
        },
        body: JSON.stringify({ con: data.con, validity: data.validity }),
      },
      result: {
        url: (o: { url: string }, { sessionToken }: { sessionToken: string }) =>
          `${o.url}/v2/request/jwt/${sessionToken}`,
        headers: data.header,
        parseResponse: (r: Response) => r.text(),
      },
    },
  });

  yivi.use(YiviClient);
  yivi.use(YiviWeb);
  yivi
    .start()
    .then((jwt: string) => {
      // Send JWT back to the parent window
      Office.context.ui.messageParent(JSON.stringify({ jwt }));
    })
    .catch((e: Error) => {
      console.error("[PostGuard Dialog] Yivi error:", e);
      Office.context.ui.messageParent(JSON.stringify({ error: e.message || "Yivi authentication failed" }));
    });
}

Office.onReady(() => {
  initializeDialog();
});
