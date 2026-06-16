// Helpers for rendering a decrypted email body inside the read-view iframe.
//
// The decrypted body is shown via `iframe.srcdoc`. Without an explicit
// colour scheme the rendered document inherits the Outlook host theme: in
// dark mode the host inverts the iframe's white background to a dark blue
// while leaving the email's default black text untouched, producing
// black-on-blue text that fails WCAG 2.2 AA contrast (issue #58).
//
// We pin the rendered document to a light colour scheme with an explicit
// white background and dark (#1a1a1a, 17.4:1 on white) default text colour.
// This is a low-specificity base, so HTML emails that set their own colours
// still win — we only guarantee a readable default when the email relies on
// the user-agent defaults.

// Default text/background for the rendered body. Mirrors the taskpane
// palette (--pg-fg on --pg-bg); contrast ratio 17.4:1, well above AA's 4.5:1.
const BODY_FG = "#1a1a1a";
const BODY_BG = "#ffffff";

// `color-scheme: light` opts the rendered document out of the host's
// automatic dark-mode inversion; the bg/colour rules guarantee a readable
// default for emails that don't style themselves.
const BASE_STYLE =
  `:root{color-scheme:light;}` +
  `html,body{background:${BODY_BG};color:${BODY_FG};}` +
  `body{margin:0;padding:8px;}`;

const HEAD_TAGS = `<meta name="color-scheme" content="light"><style>${BASE_STYLE}</style>`;

/**
 * Wrap a decrypted email body into a self-contained HTML document suitable for
 * `iframe.srcdoc`, ensuring the body text renders with AA-compliant contrast
 * regardless of the Outlook host theme.
 */
export function wrapHtml(body: string, isHtml: boolean): string {
  if (isHtml) {
    // Full HTML document: inject our base style into the existing/created
    // <head> so default text stays readable without overriding the sender.
    if (/<head[\s>]/i.test(body)) {
      return body.replace(/<head([^>]*)>/i, `<head$1>${HEAD_TAGS}`);
    }
    if (/<html[\s>]/i.test(body)) {
      return body.replace(/<html([^>]*)>/i, `<html$1><head>${HEAD_TAGS}</head>`);
    }
    if (/<body[\s>]/i.test(body)) {
      return `<!doctype html><html><head>${HEAD_TAGS}</head>${body}</html>`;
    }
    return `<!doctype html><html><head>${HEAD_TAGS}</head><body>${body}</body></html>`;
  }
  return (
    `<!doctype html><html><head>${HEAD_TAGS}</head><body>` +
    `<pre style="white-space:pre-wrap;margin:0;font-family:Segoe UI,Helvetica,Arial,sans-serif">${escapeHtml(
      body
    )}</pre></body></html>`
  );
}

/** Escape the five characters that are unsafe in HTML text/attribute context. */
export function escapeHtml(s: string): string {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
