// Entry point for the PostGuard taskpane. Detects whether we are in compose
// or read mode and dispatches to the corresponding view.

import { t } from "../lib/i18n";
import { isComposeMode } from "../lib/office-helpers";
import { stringifyError } from "../lib/stringify-error";
import { mountComposeView } from "./compose-view";
import { mountReadView } from "./read-view";
import { mountSettingsView } from "./settings-view";

const views = {
  loading: byId("view-loading"),
  compose: byId("view-compose"),
  read_encrypted: byId("view-read-encrypted"),
  read_was_encrypted: byId("view-read-was-encrypted"),
  read_noop: byId("view-read-noop"),
  decrypted: byId("view-decrypted"),
  yivi: byId("view-yivi"),
  error: byId("view-error"),
  settings: byId("view-settings"),
};

export type ViewName = keyof typeof views;

// Views from which the Settings footer button is visible. Transient
// views (loading/yivi/error) and the settings view itself stay clean.
const SETTINGS_ENTRY_VIEWS: ReadonlySet<ViewName> = new Set<ViewName>([
  "compose",
  "read_encrypted",
  "read_was_encrypted",
  "read_noop",
  "decrypted",
]);

let lastSettingsEntryView: ViewName = "compose";

export function showView(name: ViewName): void {
  for (const [k, el] of Object.entries(views)) {
    if (el) el.hidden = k !== name;
  }
  if (SETTINGS_ENTRY_VIEWS.has(name)) {
    lastSettingsEntryView = name;
  }
  const footer = byId("pg-footer");
  if (footer) footer.hidden = !SETTINGS_ENTRY_VIEWS.has(name);
}

export function showError(message: string): void {
  const errEl = byId("pg-error-text");
  if (errEl) errEl.textContent = message;
  showView("error");
}

export function setStatus(message: string, kind: "info" | "error" = "info"): void {
  const el = byId("pg-status");
  if (!el) return;
  if (!message) {
    el.classList.add("pg-status-hidden");
    el.textContent = "";
    return;
  }
  el.classList.remove("pg-status-hidden");
  el.classList.toggle("pg-status-error", kind === "error");
  el.textContent = message;
}

function byId(id: string): HTMLElement | null {
  return document.getElementById(id);
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showError("PostGuard only runs inside Outlook.");
    return;
  }

  const retry = byId("pg-error-retry") as HTMLButtonElement | null;
  if (retry) {
    retry.textContent = t("errorRetry");
    retry.addEventListener("click", () => bootstrap());
  }

  const yiviCancel = byId("pg-btn-yivi-cancel") as HTMLButtonElement | null;
  if (yiviCancel) yiviCancel.textContent = t("yiviCancel");

  const noopText = byId("pg-read-noop-text");
  if (noopText) noopText.textContent = t("readNoopMessage");

  const settingsLabel = byId("pg-open-settings-label");
  if (settingsLabel) settingsLabel.textContent = t("settingsOpen");

  const settingsBtn = byId("pg-open-settings") as HTMLButtonElement | null;
  if (settingsBtn) {
    settingsBtn.setAttribute("aria-label", t("settingsOpen"));
    settingsBtn.title = t("settingsOpen");
    settingsBtn.addEventListener("click", () => mountSettingsView(lastSettingsEntryView));
  }

  bootstrap();
});

async function bootstrap(): Promise<void> {
  showView("loading");
  setStatus("");
  try {
    const compose = isComposeMode();
    const subjType = typeof (Office.context.mailbox.item as { subject?: unknown })?.subject;

    console.log(
      `[pg-taskpane] bootstrap platform=${Office.context.platform} ` +
        `host=${Office.context.host} compose=${compose} subjectType=${subjType}`
    );
    if (compose) {
      await mountComposeView();
    } else {
      await mountReadView();
    }
  } catch (err) {
    const detail = stringifyError(err);
    const message = detail || "PostGuard failed to start.";

    console.error(`[pg-taskpane] bootstrap threw: ${message}`, err);
    showError(message);
  }
}
