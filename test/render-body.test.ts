// Tests for the decrypted-body rendering helpers (src/lib/render-body.ts).
//
// Run with Node's built-in test runner (no extra dependencies):
//   npm test
//
// These pin the WCAG 2.2 AA contrast fix from issue #58: the rendered email
// body must declare a light colour scheme with an explicit white background
// and dark default text so the Outlook host's dark mode cannot invert it into
// unreadable black-on-blue.

import { test } from "node:test";
import assert from "node:assert/strict";
import { wrapHtml, escapeHtml } from "../src/lib/render-body.ts";

// Relative luminance + contrast ratio per WCAG 2.x definitions.
function luminance(hex: string): number {
  const h = hex.replace("#", "");
  const channels = [0, 2, 4].map((i) => {
    const c = parseInt(h.substr(i, 2), 16) / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}
function contrast(a: string, b: string): number {
  const l1 = luminance(a);
  const l2 = luminance(b);
  return (Math.max(l1, l2) + 0.05) / (Math.min(l1, l2) + 0.05);
}

test("default body colours meet WCAG 2.2 AA contrast (>= 4.5:1)", () => {
  // #1a1a1a on #ffffff — the pinned defaults in render-body.ts.
  assert.ok(contrast("#1a1a1a", "#ffffff") >= 4.5);
});

test("plaintext body opts out of host dark-mode inversion", () => {
  const out = wrapHtml("hello world", false);
  assert.match(out, /color-scheme:\s*light/);
  assert.match(out, /<meta name="color-scheme" content="light">/);
  // Explicit readable background + text so defaults never rely on the host.
  assert.match(out, /background:#ffffff/);
  assert.match(out, /color:#1a1a1a/);
  assert.match(out, /hello world/);
});

test("plaintext body is HTML-escaped", () => {
  const out = wrapHtml('<script>alert("x")</script>', false);
  assert.match(out, /&lt;script&gt;/);
  assert.doesNotMatch(out, /<script>alert/);
});

test("plain HTML fragment gets a head with the contrast base style", () => {
  const out = wrapHtml("<p>hi</p>", true);
  assert.match(out, /<head>.*color-scheme:\s*light.*<\/head>/s);
  assert.match(out, /<p>hi<\/p>/);
});

test("full HTML document with a head gets the base style injected", () => {
  const out = wrapHtml("<html><head><title>t</title></head><body><p>hi</p></body></html>", true);
  // Original content preserved.
  assert.match(out, /<title>t<\/title>/);
  assert.match(out, /<p>hi<\/p>/);
  // Our colour scheme injected into the existing head exactly once.
  assert.equal(out.match(/color-scheme:\s*light/g)?.length, 1);
});

test("HTML document without a head still gets one", () => {
  const out = wrapHtml("<html><body><p>hi</p></body></html>", true);
  assert.match(out, /<head>.*color-scheme:\s*light.*<\/head>/s);
  assert.match(out, /<p>hi<\/p>/);
});

test("escapeHtml escapes the unsafe characters", () => {
  assert.equal(escapeHtml('<a href="x">&'), "&lt;a href=&quot;x&quot;&gt;&amp;");
});

// The decrypted body may contain hostile markup (remote <img> tracking
// pixels, remote CSS, etc.). The wrapper must inject a restrictive CSP meta
// so the rendered document cannot fetch remote resources regardless of what
// the email markup asks for. Without it, opening a message leaks the reader's
// IP/activity to arbitrary origins.
function csp(out: string): string {
  const m = out.match(
    /<meta http-equiv="Content-Security-Policy" content="([^"]*)">/i
  );
  assert.ok(m, "CSP meta tag must be present");
  return m![1];
}

test("plaintext body carries a restrictive CSP meta tag", () => {
  const policy = csp(wrapHtml("hello", false));
  assert.match(policy, /default-src 'none'/);
  // Only inlined images/fonts; no remote fetches. No http(s) source anywhere.
  assert.match(policy, /img-src data: blob:/);
  assert.match(policy, /font-src data:/);
  assert.doesNotMatch(policy, /https?:/);
});

test("CSP allows inline styles the wrapper and emails rely on", () => {
  const policy = csp(wrapHtml("<p>hi</p>", true));
  assert.match(policy, /style-src 'unsafe-inline'/);
});

for (const [label, body] of [
  ["plain fragment", "<p>hi</p>"],
  ["existing head", "<html><head><title>t</title></head><body>hi</body></html>"],
  ["no head", "<html><body>hi</body></html>"],
  ["body only", "<body>hi</body>"],
] as const) {
  test(`CSP is injected exactly once (${label})`, () => {
    const out = wrapHtml(body, true);
    assert.equal(
      out.match(/http-equiv="Content-Security-Policy"/gi)?.length,
      1,
      "exactly one CSP meta tag"
    );
    // And it sits inside the document head.
    assert.match(out, /<head>.*Content-Security-Policy.*<\/head>/is);
  });
}

test("CSP precedes the body content so it governs the whole document", () => {
  const out = wrapHtml('<img src="https://tracker.example/pixel.gif">', true);
  const cspAt = out.search(/Content-Security-Policy/i);
  const imgAt = out.search(/tracker\.example/i);
  assert.ok(cspAt >= 0 && imgAt >= 0);
  assert.ok(cspAt < imgAt, "CSP meta must appear before body markup");
});
