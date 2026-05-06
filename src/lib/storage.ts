// Persistent per-mailbox storage using Office roamingSettings.
// roamingSettings is JSON-serializable, ~32KB total budget.

/* global Office */

export function getSetting<T>(key: string, fallback: T): T {
  const v = Office.context.roamingSettings.get(key) as T | undefined;
  return v === undefined || v === null ? fallback : v;
}

export function setSetting<T>(key: string, value: T): Promise<void> {
  Office.context.roamingSettings.set(key, value);
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

export function removeSetting(key: string): Promise<void> {
  Office.context.roamingSettings.remove(key);
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}
