// Settings view: sender-attribute prefills (mandatory disclosure values
// for fullname / DOB / mobile) and the advanced "skip the open-a-dialog
// confirmation" toggle. All values are persisted to roamingSettings so
// the OnMessageSend launchevent runtime can read them too.
//
// Listeners are attached exactly once on first mount (tracked via the
// `wired` flag). Subsequent mounts only refresh the input values from
// roamingSettings — no DOM cloning, which avoids the browser-specific
// cloneNode-with-input.value quirks the previous implementation hit.

import {
  SIGN_PREFILL_FULLNAME,
  SIGN_PREFILL_DATEOFBIRTH,
  SIGN_PREFILL_MOBILE,
  SignPrefillType,
  getAllowOptimisticDialog,
  getEncryptionEnabled,
  getSignPrefills,
  setAllowOptimisticDialog,
  setEncryptionEnabled,
  setSignPrefills,
} from "../lib/settings";
import { byId } from "../lib/dom";
import { t } from "../lib/i18n";
import { showView, setStatus, ViewName } from "./taskpane";

let wired = false;
let returnView: ViewName = "compose";

export function mountSettingsView(returnTo: ViewName): void {
  returnView = returnTo;
  refreshLabels();
  refreshValues();
  if (!wired) {
    wireListeners();
    wired = true;
  }
  showView("settings");
}

function refreshLabels(): void {
  byId<HTMLElement>("pg-settings-title").textContent = t("settingsTitle");
  byId<HTMLElement>("pg-settings-encryption-title").textContent = t("settingsEncryptionTitle");
  byId<HTMLElement>("pg-settings-encryption-default-label").textContent = t(
    "settingsEncryptionDefaultLabel"
  );
  byId<HTMLElement>("pg-settings-encryption-default-help").textContent = t(
    "settingsEncryptionDefaultHelp"
  );
  byId<HTMLElement>("pg-settings-prefill-title").textContent = t("settingsPrefillTitle");
  byId<HTMLElement>("pg-settings-prefill-help").textContent = t("settingsPrefillHelp");
  byId<HTMLElement>("pg-settings-advanced-title").textContent = t("settingsAdvancedTitle");
  byId<HTMLElement>("pg-settings-optimistic-label").textContent = t(
    "settingsAllowOptimisticDialogLabel"
  );
  byId<HTMLElement>("pg-settings-optimistic-help").textContent = t(
    "settingsAllowOptimisticDialogHelp"
  );
  byId<HTMLElement>("pg-prefill-fullname-label").textContent = t(SIGN_PREFILL_FULLNAME);
  byId<HTMLElement>("pg-prefill-dateofbirth-label").textContent = t(SIGN_PREFILL_DATEOFBIRTH);
  byId<HTMLElement>("pg-prefill-mobile-label").textContent = t(SIGN_PREFILL_MOBILE);
  byId<HTMLButtonElement>("pg-settings-save").textContent = t("settingsSave");
  byId<HTMLButtonElement>("pg-settings-back").textContent = t("settingsBack");
}

function refreshValues(): void {
  const prefills = getSignPrefills();
  byId<HTMLInputElement>("pg-toggle-encryption-default").checked = getEncryptionEnabled();
  byId<HTMLInputElement>("pg-prefill-fullname").value = prefills[SIGN_PREFILL_FULLNAME] ?? "";
  byId<HTMLInputElement>("pg-prefill-dateofbirth").value = ddmmyyyyToHtml(
    prefills[SIGN_PREFILL_DATEOFBIRTH] ?? ""
  );
  byId<HTMLInputElement>("pg-prefill-mobile").value = prefills[SIGN_PREFILL_MOBILE] ?? "";
  byId<HTMLInputElement>("pg-toggle-allow-optimistic-dialog").checked = getAllowOptimisticDialog();
  byId<HTMLButtonElement>("pg-settings-save").disabled = false;
}

function wireListeners(): void {
  byId<HTMLInputElement>("pg-toggle-encryption-default").addEventListener("change", () => {
    const checked = byId<HTMLInputElement>("pg-toggle-encryption-default").checked;
    void setEncryptionEnabled(checked).catch((err) => {
      console.error("[pg-settings] failed to persist encryption default", err);
      setStatus(t("settingsSaveError"), "error");
    });
  });

  byId<HTMLInputElement>("pg-toggle-allow-optimistic-dialog").addEventListener("change", () => {
    const checked = byId<HTMLInputElement>("pg-toggle-allow-optimistic-dialog").checked;
    void setAllowOptimisticDialog(checked).catch((err) => {
      console.error("[pg-settings] failed to persist toggle", err);
      setStatus(t("settingsSaveError"), "error");
    });
  });

  byId<HTMLButtonElement>("pg-settings-save").addEventListener("click", () => {
    const save = byId<HTMLButtonElement>("pg-settings-save");
    save.disabled = true;
    const next: Partial<Record<SignPrefillType, string>> = {
      [SIGN_PREFILL_FULLNAME]: byId<HTMLInputElement>("pg-prefill-fullname").value,
      [SIGN_PREFILL_DATEOFBIRTH]: htmlToDdmmyyyy(
        byId<HTMLInputElement>("pg-prefill-dateofbirth").value
      ),
      [SIGN_PREFILL_MOBILE]: byId<HTMLInputElement>("pg-prefill-mobile").value,
    };
    console.log("[pg-settings] saving prefills:", next);
    setSignPrefills(next)
      .then(() => {
        console.log("[pg-settings] persisted; readback:", getSignPrefills());
        setStatus(t("settingsSaved"));
        setTimeout(() => setStatus(""), 2000);
        showView(returnView);
      })
      .catch((err) => {
        console.error("[pg-settings] failed to persist prefills", err);
        setStatus(t("settingsSaveError"), "error");
        save.disabled = false;
      });
  });

  byId<HTMLButtonElement>("pg-settings-back").addEventListener("click", () => {
    setStatus("");
    showView(returnView);
  });
}

function ddmmyyyyToHtml(ddmmyyyy: string): string {
  if (!ddmmyyyy) return "";
  const p = ddmmyyyy.split("-");
  return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : "";
}

function htmlToDdmmyyyy(yyyymmdd: string): string {
  if (!yyyymmdd) return "";
  const p = yyyymmdd.split("-");
  return p.length === 3 ? `${p[2]}-${p[1]}-${p[0]}` : "";
}
