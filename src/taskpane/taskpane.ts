// Entry point for the PostGuard taskpane. Detects whether we are in compose
// or read mode and dispatches to the corresponding view.

import { isComposeMode } from "../lib/office-helpers";
import { mountComposeView } from "./compose-view";
import { mountReadView } from "./read-view";

/* global Office */

const views = {
  loading: byId("view-loading"),
  compose: byId("view-compose"),
  read_encrypted: byId("view-read-encrypted"),
  read_was_encrypted: byId("view-read-was-encrypted"),
  read_noop: byId("view-read-noop"),
  decrypted: byId("view-decrypted"),
  yivi: byId("view-yivi"),
  error: byId("view-error"),
};

export type ViewName = keyof typeof views;

export function showView(name: ViewName): void {
  for (const [k, el] of Object.entries(views)) {
    if (el) el.hidden = k !== name;
  }
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
  if (retry) retry.addEventListener("click", () => bootstrap());

  bootstrap();
});

async function bootstrap(): Promise<void> {
  showView("loading");
  setStatus("");
  try {
    if (isComposeMode()) {
      await mountComposeView();
    } else {
      await mountReadView();
    }
  } catch (err) {
    const message = err instanceof Error ? err.message : "PostGuard failed to start.";
    showError(message);
  }
}
