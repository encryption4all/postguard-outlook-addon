// Function file required by the manifest's <FunctionFile> entry. The
// PostGuard add-in does not register any ExecuteFunction commands today —
// every action goes through the taskpane — but the file must exist and
// load Office.js so the runtime can mount it.

/* global Office */

Office.onReady(() => {
  // No-op: kept to satisfy the manifest's FunctionFile reference.
});
