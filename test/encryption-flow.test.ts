import {
  runEncryptionFlow,
  FailureDeps,
  DialogMessageLike,
} from "../src/yivi-dialog/encryption-flow";

// A recording set of dependencies. Each side-effect pushes a tag so we can
// assert ordering, and any of them can be made to throw to simulate a stale
// DOM node or a dead parent window.
function makeDeps(
  overrides: Partial<FailureDeps & { runEncryption: () => Promise<DialogMessageLike> }> = {}
) {
  const calls: string[] = [];
  const posted: DialogMessageLike[] = [];
  const errors: string[] = [];
  const completed: { message: string; isError?: boolean }[] = [];
  const deps = {
    runEncryption:
      overrides.runEncryption ??
      (async () => ({ type: "encrypt-result", subject: "s" }) as DialogMessageLike),
    postChunkedToParent:
      overrides.postChunkedToParent ??
      ((p: DialogMessageLike) => {
        calls.push("post");
        posted.push(p);
      }),
    showError:
      overrides.showError ??
      ((m: string) => {
        calls.push("showError");
        errors.push(m);
      }),
    showCompleted:
      overrides.showCompleted ??
      ((message: string, isError?: boolean) => {
        calls.push("showCompleted");
        completed.push({ message, isError });
      }),
    log: overrides.log ?? (() => undefined),
    stringifyError:
      overrides.stringifyError ?? ((e: unknown) => (e instanceof Error ? e.message : String(e))),
    isUploadSessionExpired: overrides.isUploadSessionExpired ?? (() => false),
  };
  return { deps, calls, posted, errors, completed };
}

describe("runEncryptionFlow", () => {
  it("posts the result and shows completion on success", async () => {
    const { deps, posted, completed, errors } = makeDeps();
    await runEncryptionFlow({}, deps);

    expect(posted).toEqual([{ type: "encrypt-result", subject: "s" }]);
    expect(completed).toEqual([
      { message: "Encrypted and sent. You can close this window.", isError: undefined },
    ]);
    expect(errors).toEqual([]);
  });

  it("reports a generic failure when runEncryption rejects", async () => {
    const { deps, posted, errors, completed } = makeDeps({
      runEncryption: async () => {
        throw new Error("boom");
      },
    });
    await runEncryptionFlow({}, deps);

    expect(errors).toEqual(["Encryption failed: boom"]);
    expect(posted).toEqual([{ type: "encrypt-error", message: "boom" }]);
    expect(completed).toEqual([{ message: "boom", isError: true }]);
  });

  it("surfaces the expired-session message and code distinctly", async () => {
    const { deps, posted, errors } = makeDeps({
      runEncryption: async () => {
        throw new Error("404");
      },
      isUploadSessionExpired: () => true,
    });
    await runEncryptionFlow({}, deps);

    const msg = "The upload session expired. Please start a new send.";
    expect(errors).toEqual([msg]);
    expect(posted).toEqual([
      { type: "encrypt-error", message: msg, code: "upload_session_expired" },
    ]);
  });

  // The core of issue #78: a throw in a success-arm side-effect must funnel
  // into the failure path, not leak as an unhandled rejection.
  it("falls back to the failure path when postChunkedToParent throws on success", async () => {
    let firstPost = true;
    const errors: string[] = [];
    const posted: DialogMessageLike[] = [];
    const { deps } = makeDeps({
      postChunkedToParent: (p: DialogMessageLike) => {
        if (firstPost) {
          firstPost = false;
          throw new Error("parent window gone");
        }
        posted.push(p);
      },
      showError: (m: string) => errors.push(m),
    });

    await expect(runEncryptionFlow({}, deps)).resolves.toBeUndefined();
    // The failure reporter ran: an error was shown and the error payload was
    // posted (the retry succeeds because firstPost is now false).
    expect(errors).toEqual(["Encryption failed: parent window gone"]);
    expect(posted).toEqual([{ type: "encrypt-error", message: "parent window gone" }]);
  });

  it("never rejects even when every failure-path side-effect throws", async () => {
    const { deps } = makeDeps({
      runEncryption: async () => {
        throw new Error("boom");
      },
      showError: () => {
        throw new Error("stale node");
      },
      postChunkedToParent: () => {
        throw new Error("dead parent");
      },
      showCompleted: () => {
        throw new Error("dead dom");
      },
    });

    await expect(runEncryptionFlow({}, deps)).resolves.toBeUndefined();
  });

  it("still posts the parent notification when showError throws", async () => {
    const posted: DialogMessageLike[] = [];
    const completed: { message: string; isError?: boolean }[] = [];
    const { deps } = makeDeps({
      runEncryption: async () => {
        throw new Error("boom");
      },
      showError: () => {
        throw new Error("stale node");
      },
      postChunkedToParent: (p: DialogMessageLike) => posted.push(p),
      showCompleted: (message: string, isError?: boolean) => completed.push({ message, isError }),
    });

    await runEncryptionFlow({}, deps);
    // showError throwing must not stop the encrypt-error reaching the parent
    // (it drives the Smart Alert) nor the dialog moving to its error state.
    expect(posted).toEqual([{ type: "encrypt-error", message: "boom" }]);
    expect(completed).toEqual([{ message: "boom", isError: true }]);
  });

  it("falls back to a generic message when stringifyError itself throws", async () => {
    const posted: DialogMessageLike[] = [];
    const errors: string[] = [];
    const { deps } = makeDeps({
      runEncryption: async () => {
        throw new Error("boom");
      },
      stringifyError: () => {
        throw new Error("cannot stringify");
      },
      showError: (m: string) => errors.push(m),
      postChunkedToParent: (p: DialogMessageLike) => posted.push(p),
    });

    await runEncryptionFlow({}, deps);
    const generic = "Encryption failed. Please close this window and try again.";
    expect(errors).toEqual([generic]);
    expect(posted).toEqual([{ type: "encrypt-error", message: generic }]);
  });
});
