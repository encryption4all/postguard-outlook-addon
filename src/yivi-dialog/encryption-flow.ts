// Success / failure orchestration for the Yivi dialog's encryption run.
//
// The success and failure handlers both call DOM + Office.js helpers
// (showError, postChunkedToParent, showCompleted) that can throw at
// runtime — e.g. messageParent fails because the parent window is gone, or
// a helper touches a stale DOM node. Previously these handlers were the
// success/failure arms of `runEncryption(req).then(onOk, onErr)`: a throw
// inside either arm produced a rejected promise with no handler, surfacing
// only as an unhandled-rejection log while the dialog stayed on its current
// view with no error shown to the user.
//
// runEncryptionFlow() guarantees a failure path is always reachable:
//   - a throw anywhere in the success arm (including postChunkedToParent /
//     showCompleted) is funnelled into the failure reporter, so the user
//     still sees "Encryption failed";
//   - inside the failure reporter every side-effect is attempted
//     independently and best-effort, so a throw in one (e.g. showError on a
//     stale node) cannot stop the parent window from receiving the
//     encrypt-error notification — and none of them may leak.
//
// The flow is dependency-injected (no direct Office.js / pg-js imports) so
// it can be unit-tested in plain Node.

export interface DialogMessageLike {
  type: string;
  [key: string]: unknown;
}

/** The side-effecting helpers the failure reporter needs. */
export interface FailureDeps {
  postChunkedToParent: (payload: DialogMessageLike) => void;
  showError: (message: string) => void;
  showCompleted: (message: string, isError?: boolean) => void;
  log: (msg: string) => void;
  stringifyError: (err: unknown) => string;
  /** True when the rejection is pg-js's UploadSessionExpiredError. */
  isUploadSessionExpired: (err: unknown) => boolean;
}

export interface EncryptionFlowDeps<TReq, TResult> extends FailureDeps {
  runEncryption: (req: TReq) => Promise<TResult>;
}

const GENERIC_FAILURE = "Encryption failed. Please close this window and try again.";

/** Run `fn`, swallowing any throw. Used so a failing side-effect in the
 *  error path cannot stop the others or escape as an unhandled rejection. */
function safe(fn: () => void): void {
  try {
    fn();
  } catch {
    // Nothing left to do — swallow so the guard itself can never throw.
  }
}

function reportFailure(err: unknown, deps: FailureDeps): void {
  let message: string;
  let errorText: string; // what showError displays in the dialog
  let payload: DialogMessageLike; // what we post back to the parent window
  try {
    const expired = deps.isUploadSessionExpired(err);
    // pg-js raises UploadSessionExpiredError when cryptify's structured
    // 404 says the upload session is gone (TTL expired, server restart,
    // unknown UUID, or wrong recovery_token). Surface that distinctly so
    // the Smart Alert can tell the user "start a new send" instead of a
    // generic "encryption failed" — see issue #82.
    message = expired
      ? "The upload session expired. Please start a new send."
      : deps.stringifyError(err);
    errorText = expired ? message : `Encryption failed: ${message}`;
    payload = expired
      ? { type: "encrypt-error", message, code: "upload_session_expired" }
      : { type: "encrypt-error", message };
    deps.log(
      `encryption failed${expired ? " (upload session expired)" : ""}: ${deps.stringifyError(err)}`
    );
  } catch {
    // Deriving the message itself threw (e.g. a custom isUploadSessionExpired
    // or stringifyError choked on an exotic value). Fall back to a generic,
    // non-throwing payload so the user still gets a failure surface.
    message = GENERIC_FAILURE;
    errorText = GENERIC_FAILURE;
    payload = { type: "encrypt-error", message };
    safe(() => deps.log("encryption failed; could not derive error detail"));
  }
  // Each side-effect is best-effort and independent: posting the
  // encrypt-error to the parent (which drives the Smart Alert) must not be
  // skipped just because showError threw on a stale DOM node, and no throw
  // here may surface as an unhandled rejection.
  safe(() => deps.showError(errorText));
  safe(() => deps.postChunkedToParent(payload));
  safe(() => deps.showCompleted(message, true));
}

/**
 * Run the encryption and report the result, never leaking a rejection.
 *
 * Resolves once the dialog has been moved into its completed (success or
 * error) state. It is designed never to reject — every throwing path is
 * funnelled into {@link reportFailure}, which itself never throws — but
 * callers should still guard the boundary defensively.
 */
export async function runEncryptionFlow<TReq, TResult>(
  req: TReq,
  deps: EncryptionFlowDeps<TReq, TResult>
): Promise<void> {
  try {
    const result = await deps.runEncryption(req);
    deps.log("encryption complete; posting result");
    deps.postChunkedToParent(result as unknown as DialogMessageLike);
    deps.showCompleted("Encrypted and sent. You can close this window.");
  } catch (err) {
    reportFailure(err, deps);
  }
}
