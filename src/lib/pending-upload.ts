// Persist the `{uuid, recoveryToken}` pair pg-js's `onUploadInit`
// callback hands out, so a future session can probe cryptify for the
// status of an interrupted upload via `resumeUpload`. See issue #82.
//
// Storage backends:
//  - `roaming`   — `Office.context.roamingSettings`, per-mailbox. Used
//                  by the taskpane compose view.
//  - `local`     — `window.localStorage`. Used by the yivi-dialog
//                  webview, which is its own origin and does not
//                  reliably expose roamingSettings.
//
// Schema is a single-slot object; we don't track concurrent uploads.
// The addon's UX is one send at a time, so a second upload starting
// overwrites a still-pending first one — acceptable for now.

import { resumeUpload, UploadSessionExpiredError } from "@e4a/pg-js";

export interface PendingUpload {
  uuid: string;
  recoveryToken: string;
  savedAt: number;
}

export type Backend = "roaming" | "local";

const KEY = "pg.pendingUpload";

function readRoaming(): PendingUpload | null {
  const v = Office.context.roamingSettings.get(KEY) as PendingUpload | undefined;
  return v ?? null;
}

function writeRoaming(value: PendingUpload | null): Promise<void> {
  if (value === null) Office.context.roamingSettings.remove(KEY);
  else Office.context.roamingSettings.set(KEY, value);
  return new Promise((resolve) => {
    Office.context.roamingSettings.saveAsync(() => resolve());
  });
}

function readLocal(): PendingUpload | null {
  try {
    const raw = window.localStorage.getItem(KEY);
    if (!raw) return null;
    return JSON.parse(raw) as PendingUpload;
  } catch {
    return null;
  }
}

function writeLocal(value: PendingUpload | null): void {
  try {
    if (value === null) window.localStorage.removeItem(KEY);
    else window.localStorage.setItem(KEY, JSON.stringify(value));
  } catch {
    // localStorage may be unavailable (private mode); silently drop.
  }
}

export function recordPendingUpload(
  backend: Backend,
  info: { uuid: string; recoveryToken: string }
): void {
  const entry: PendingUpload = { ...info, savedAt: Date.now() };
  if (backend === "roaming") void writeRoaming(entry);
  else writeLocal(entry);
}

export function clearPendingUpload(backend: Backend): void {
  if (backend === "roaming") void writeRoaming(null);
  else writeLocal(null);
}

export function readPendingUpload(backend: Backend): PendingUpload | null {
  return backend === "roaming" ? readRoaming() : readLocal();
}

// Probe cryptify for a stale entry's status. Returns `uploaded` bytes
// on success, null when the session is gone or the probe failed.
// Always clears the entry afterwards: pg-js 1.8.0 does not yet accept
// a pre-resumed FileState back into `createEnvelope`, so the addon
// cannot continue the encryption from where it stopped — the probe is
// diagnostic and the entry has no other use.
export async function probeAndClearPendingUpload(
  backend: Backend,
  cryptifyUrl: string,
  minAgeMs = 60_000
): Promise<number | null> {
  const entry = readPendingUpload(backend);
  if (!entry) return null;
  if (Date.now() - entry.savedAt < minAgeMs) return null;
  try {
    const { uploaded } = await resumeUpload(cryptifyUrl, entry.uuid, entry.recoveryToken);
    clearPendingUpload(backend);
    return uploaded;
  } catch (err) {
    if (err instanceof UploadSessionExpiredError) {
      clearPendingUpload(backend);
      return null;
    }
    return null;
  }
}
