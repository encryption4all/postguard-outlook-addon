// Typed wrappers around the roaming-settings keys used by the taskpane and
// the OnMessageSend launchevent runtime. Keeping the key + default in one
// place stops the two readers from drifting (taskpane writes, launchevent
// reads).

import { getSetting, setSetting } from "./storage";
import { AttributeRequest } from "./types";

// When true, the launchevent skips Office's "PostGuard wants to open a
// dialog" confirmation and tries to open the Yivi dialog directly. Only
// works in browsers/hosts where the user has already granted popup
// permission to the add-in's origin (Safari: Settings → Websites →
// Pop-ups → Allow); elsewhere the optimistic attempt is blocked and the
// launchevent falls back to the prompted open once.
export const ALLOW_OPTIMISTIC_DIALOG_KEY = "pg.allowOptimisticDialog";

export function getAllowOptimisticDialog(): boolean {
  return getSetting<boolean>(ALLOW_OPTIMISTIC_DIALOG_KEY, false);
}

export function setAllowOptimisticDialog(value: boolean): Promise<void> {
  return setSetting<boolean>(ALLOW_OPTIMISTIC_DIALOG_KEY, value);
}

// Sign-attribute prefills. The three Yivi attribute types listed below are
// always offered for disclosure when the user signs a message. If the user
// has filled in a value in Settings, it is sent as a mandatory disclosure
// (the Yivi app must match it). If the value is blank, the attribute is
// sent as `optional: true` — the user discloses it in the Yivi app or
// skips it.
export const SIGN_PREFILL_FULLNAME = "pbdf.gemeente.personalData.fullname";
export const SIGN_PREFILL_DATEOFBIRTH = "pbdf.gemeente.personalData.dateofbirth";
export const SIGN_PREFILL_MOBILE = "pbdf.sidn-pbdf.mobilenumber.mobilenumber";
export const SIGN_PREFILL_TYPES = [
  SIGN_PREFILL_FULLNAME,
  SIGN_PREFILL_DATEOFBIRTH,
  SIGN_PREFILL_MOBILE,
] as const;
export type SignPrefillType = (typeof SIGN_PREFILL_TYPES)[number];

const SIGN_PREFILLS_KEY = "pg.signPrefills";

export function getSignPrefills(): Partial<Record<SignPrefillType, string>> {
  const raw = getSetting<Partial<Record<SignPrefillType, string>>>(SIGN_PREFILLS_KEY, {});
  // Defensive copy — roamingSettings returns the live object on subsequent
  // reads and callers shouldn't mutate persisted state by accident.
  return { ...raw };
}

export function setSignPrefills(values: Partial<Record<SignPrefillType, string>>): Promise<void> {
  // Persist only the canonical keys; drop empty strings so a never-filled
  // attribute stays "optional" rather than "mandatory with empty value".
  const cleaned: Partial<Record<SignPrefillType, string>> = {};
  for (const k of SIGN_PREFILL_TYPES) {
    const v = (values[k] ?? "").trim();
    if (v) cleaned[k] = v;
  }
  return setSetting(SIGN_PREFILLS_KEY, cleaned);
}

// Builds the attribute list passed to pg.sign.yivi({ attributes }). Each of
// SIGN_PREFILL_TYPES is included exactly once: as a mandatory { t, v } when
// the user prefilled a value, otherwise as { t, optional: true }.
export function buildSignAttributes(): AttributeRequest[] {
  const prefills = getSignPrefills();
  return SIGN_PREFILL_TYPES.map((t) => {
    const v = (prefills[t] ?? "").trim();
    return v ? { t, v } : { t, optional: true };
  });
}
