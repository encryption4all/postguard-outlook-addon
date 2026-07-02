// Best-effort coercion of an unknown thrown / rejected value into a
// human-readable string. Outlook for Mac's WKWebView surfaces some
// failures as plain object rejections rather than Error instances, and
// Office.AsyncResult.error has shape `{ code, name, message }`. A naive
// `String(err)` collapses those to `"[object Object]"`, which loses
// every diagnostic clue. This helper preserves whatever shape we got.
export function stringifyError(err: unknown): string {
  if (err instanceof Error) {
    // Never fold the stack into the returned string — it can leak internal
    // file paths and implementation details into user-facing UI (Smart Alert
    // dialog / taskpane). Log it to the console for diagnostics instead and
    // return only the human-readable message. See GHSA-8rxw-3qj6-p59v.
    if (err.stack) console.error(err.stack);
    return err.message;
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
