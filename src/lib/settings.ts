// Typed wrappers around the roaming-settings keys used by the taskpane and
// the OnMessageSend launchevent runtime. Keeping the key + default in one
// place stops the two readers from drifting (taskpane writes, launchevent
// reads).

import { getSetting, setSetting } from "./storage";

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
