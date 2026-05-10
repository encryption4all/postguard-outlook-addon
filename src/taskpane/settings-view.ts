// Settings view: sender-attribute prefills (mandatory disclosure values
// for fullname / DOB / mobile) and the advanced "skip the open-a-dialog
// confirmation" toggle. All values are persisted to roamingSettings so
// the OnMessageSend launchevent runtime can read them too.
//
// Prefills require an explicit Save click — auto-saving on `change` (blur)
// lost data when the user clicked Back without blurring the field first.
// The toggle still saves immediately because it's a single discrete click.

import {
  SIGN_PREFILL_FULLNAME,
  SIGN_PREFILL_DATEOFBIRTH,
  SIGN_PREFILL_MOBILE,
  SignPrefillType,
  getAllowOptimisticDialog,
  getSignPrefills,
  setAllowOptimisticDialog,
  setSignPrefills,
} from "../lib/settings";
import { t } from "../lib/i18n";
import { showView, setStatus, ViewName } from "./taskpane";

export function mountSettingsView(returnTo: ViewName): void {
  byId<HTMLElement>("pg-settings-title").textContent = t("settingsTitle");
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

  const prefills = getSignPrefills();
  // DOB is stored as DD-MM-YYYY (Yivi's format); <input type="date"> uses
  // YYYY-MM-DD. Round-trip through the same helpers the policy editor
  // uses for the Manage Access date input.
  const fullname = freshInput("pg-prefill-fullname", prefills[SIGN_PREFILL_FULLNAME] ?? "");
  const dob = freshInput(
    "pg-prefill-dateofbirth",
    ddmmyyyyToHtml(prefills[SIGN_PREFILL_DATEOFBIRTH] ?? "")
  );
  const mobile = freshInput("pg-prefill-mobile", prefills[SIGN_PREFILL_MOBILE] ?? "");

  // The advanced toggle is a single discrete click; persisting immediately
  // keeps the UX simple and the value won't be lost on Back.
  const toggle = freshCheckbox("pg-toggle-allow-optimistic-dialog", getAllowOptimisticDialog());
  toggle.addEventListener("change", () => {
    void setAllowOptimisticDialog(toggle.checked).catch((err) => {
      console.error("[pg-settings] failed to persist toggle", err);
      setStatus(t("settingsSaveError"), "error");
    });
  });

  const save = freshButton("pg-settings-save", t("settingsSave"));
  save.addEventListener("click", () => {
    save.disabled = true;
    const next: Partial<Record<SignPrefillType, string>> = {
      [SIGN_PREFILL_FULLNAME]: fullname.value,
      [SIGN_PREFILL_DATEOFBIRTH]: htmlToDdmmyyyy(dob.value),
      [SIGN_PREFILL_MOBILE]: mobile.value,
    };
    setSignPrefills(next)
      .then(() => {
        setStatus(t("settingsSaved"));
        setTimeout(() => setStatus(""), 2000);
      })
      .catch((err) => {
        console.error("[pg-settings] failed to persist prefills", err);
        setStatus(t("settingsSaveError"), "error");
      })
      .finally(() => {
        save.disabled = false;
      });
  });

  const back = freshButton("pg-settings-back", t("settingsBack"));
  back.addEventListener("click", () => {
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

// Clone-replace so listeners attached during a previous mountSettingsView
// call are dropped before we add new ones — the view is re-mountable from
// the footer Settings button at any time.
function freshInput(id: string, value: string): HTMLInputElement {
  const stale = byId<HTMLInputElement>(id);
  const fresh = stale.cloneNode(true) as HTMLInputElement;
  fresh.value = value;
  stale.replaceWith(fresh);
  return fresh;
}

function freshCheckbox(id: string, checked: boolean): HTMLInputElement {
  const stale = byId<HTMLInputElement>(id);
  const fresh = stale.cloneNode(true) as HTMLInputElement;
  fresh.checked = checked;
  stale.replaceWith(fresh);
  return fresh;
}

function freshButton(id: string, label: string): HTMLButtonElement {
  const stale = byId<HTMLButtonElement>(id);
  const fresh = stale.cloneNode(true) as HTMLButtonElement;
  fresh.textContent = label;
  fresh.disabled = false;
  stale.replaceWith(fresh);
  return fresh;
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
