// Settings view: a single roaming-settings-backed toggle for the
// "skip the open-a-dialog confirmation" preference, with a Back button
// that returns to whichever view opened it.

import { getAllowOptimisticDialog, setAllowOptimisticDialog } from "../lib/settings";
import { t } from "../lib/i18n";
import { showView, setStatus, ViewName } from "./taskpane";

export function mountSettingsView(returnTo: ViewName): void {
  const title = byId<HTMLElement>("pg-settings-title");
  const label = byId<HTMLElement>("pg-settings-optimistic-label");
  const help = byId<HTMLElement>("pg-settings-optimistic-help");
  const toggle = byId<HTMLInputElement>("pg-toggle-allow-optimistic-dialog");
  const back = byId<HTMLButtonElement>("pg-settings-back");

  title.textContent = t("settingsTitle");
  label.textContent = t("settingsAllowOptimisticDialogLabel");
  help.textContent = t("settingsAllowOptimisticDialogHelp");
  back.textContent = t("settingsBack");

  toggle.checked = getAllowOptimisticDialog();

  // Replace any listeners from a previous mount by cloning. The settings
  // view is re-mountable from multiple entry points (compose, read), so
  // we cannot rely on a single addEventListener call per app lifetime.
  const freshToggle = toggle.cloneNode(true) as HTMLInputElement;
  toggle.replaceWith(freshToggle);
  freshToggle.addEventListener("change", () => {
    void setAllowOptimisticDialog(freshToggle.checked).catch((err) => {
      console.error("[pg-settings] failed to persist toggle", err);
      setStatus("Could not save setting. Try again.", "error");
    });
  });

  const freshBack = back.cloneNode(true) as HTMLButtonElement;
  back.replaceWith(freshBack);
  freshBack.addEventListener("click", () => {
    setStatus("");
    showView(returnTo);
  });

  showView("settings");
}

function byId<T extends HTMLElement>(id: string): T {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el as T;
}
