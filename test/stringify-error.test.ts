// Tests for the unknown→string coercion helper (src/lib/stringify-error.ts).
//
// Run with Node's built-in test runner (no extra dependencies):
//   npm test
//
// These pin the hardening from security advisory GHSA-8rxw-3qj6-p59v (#113):
// a raw stack trace must never be folded into the returned string, because
// that string is shown in the Smart Alert dialog / taskpane UI and would leak
// internal file paths and implementation details. The stack is logged to the
// console for diagnostics instead.

import { test } from "node:test";
import assert from "node:assert/strict";
import { stringifyError } from "../src/lib/stringify-error.ts";

test("returns only the message for an Error, never the stack", () => {
  const err = new Error("boom");
  // Sanity: this Error actually carries a stack in the test runtime.
  assert.ok(err.stack && err.stack.length > 0);

  const errors: unknown[] = [];
  const original = console.error;
  console.error = (...args: unknown[]) => {
    errors.push(args[0]);
  };
  try {
    const out = stringifyError(err);
    assert.equal(out, "boom");
    assert.ok(!out.includes("at "), "returned string must not contain stack frames");
    assert.ok(!out.includes("stringify-error"), "returned string must not contain file paths");
  } finally {
    console.error = original;
  }

  // The stack is still emitted to the console for diagnostics.
  assert.equal(errors.length, 1);
  assert.equal(errors[0], err.stack);
});

test("passes through a plain string unchanged", () => {
  assert.equal(stringifyError("just a string"), "just a string");
});

test("prefers a message field on a plain object rejection", () => {
  // Outlook for Mac's WKWebView / Office.AsyncResult.error shape.
  assert.equal(stringifyError({ code: 42, name: "X", message: "nope" }), "nope");
});

test("falls back to JSON for a message-less object", () => {
  assert.equal(stringifyError({ code: 7 }), JSON.stringify({ code: 7 }));
});
