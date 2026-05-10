// Persistent per-mailbox storage using Office roamingSettings.
// roamingSettings is JSON-serializable, ~32KB total budget.
//
// saveAsync can fail with code 9019 ("GenericSettingsError: An internal
// error has occurred.") when called rapidly in succession, which is what
// the Settings prefill inputs triggered on every keystroke. We serialize
// the saves through a single-flight queue and retry once on 9019.

export function getSetting<T>(key: string, fallback: T): T {
  const v = Office.context.roamingSettings.get(key) as T | undefined;
  return v === undefined || v === null ? fallback : v;
}

export function setSetting<T>(key: string, value: T): Promise<void> {
  Office.context.roamingSettings.set(key, value);
  return enqueueSave();
}

export function removeSetting(key: string): Promise<void> {
  Office.context.roamingSettings.remove(key);
  return enqueueSave();
}

let saveChain: Promise<void> = Promise.resolve();

function enqueueSave(): Promise<void> {
  const next = saveChain.catch(() => undefined).then(() => saveOnceWithRetry());
  // The chain swallows errors on subsequent links so a single failed save
  // doesn't poison the queue. The caller of setSetting / removeSetting
  // still sees the failure via the returned promise.
  saveChain = next.catch(() => undefined);
  return next;
}

function saveOnceWithRetry(attempt = 0): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }
      const code = (res.error as { code?: number } | undefined)?.code;
      if (attempt < 2 && code === 9019) {
        setTimeout(
          () => {
            saveOnceWithRetry(attempt + 1).then(resolve, reject);
          },
          150 * (attempt + 1)
        );
        return;
      }
      reject(res.error);
    });
  });
}
