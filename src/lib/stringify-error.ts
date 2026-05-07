// Best-effort coercion of an unknown thrown / rejected value into a
// human-readable string. Outlook for Mac's WKWebView surfaces some
// failures as plain object rejections rather than Error instances, and
// Office.AsyncResult.error has shape `{ code, name, message }`. A naive
// `String(err)` collapses those to `"[object Object]"`, which loses
// every diagnostic clue. This helper preserves whatever shape we got.
export function stringifyError(err: unknown): string {
  if (err instanceof Error) {
    return err.stack ? `${err.message}\n${err.stack}` : err.message;
  }
  if (typeof err === "string") return err;
  if (err && typeof err === "object") {
    const maybeMessage = (err as { message?: unknown }).message;
    if (typeof maybeMessage === "string" && maybeMessage.length > 0) {
      return maybeMessage;
    }
    try {
      return JSON.stringify(err);
    } catch {
      // fall through
    }
  }
  return String(err);
}
