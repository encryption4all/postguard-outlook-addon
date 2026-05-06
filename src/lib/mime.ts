// Minimal MIME helpers to detect PostGuard ciphertext and pull headers
// out of decrypted plaintext MIME.
//
// We deliberately do NOT try to be a full RFC 5322 parser — only the
// pieces PostGuard needs: thread-related headers and the ASCII-armored
// ciphertext block.

const ARMOR_BEGIN = "-----BEGIN POSTGUARD MESSAGE-----";
const ARMOR_END = "-----END POSTGUARD MESSAGE-----";

export function extractArmoredCiphertext(htmlOrText: string): string | null {
  if (!htmlOrText) return null;
  const begin = htmlOrText.indexOf(ARMOR_BEGIN);
  if (begin < 0) return null;
  const end = htmlOrText.indexOf(ARMOR_END, begin);
  if (end < 0) return null;
  const block = htmlOrText.slice(begin + ARMOR_BEGIN.length, end);
  // Strip whitespace and any HTML tags that may have been wrapped around it.
  return block.replace(/<[^>]+>/g, "").replace(/\s+/g, "");
}

export function looksLikePostGuard(htmlOrText: string): boolean {
  if (!htmlOrText) return false;
  return htmlOrText.indexOf(ARMOR_BEGIN) >= 0;
}

// Pull a single header value (case-insensitive) out of a raw MIME blob.
export function readMimeHeader(rawMime: string, name: string): string | undefined {
  if (!rawMime) return undefined;
  const lcName = name.toLowerCase();
  // Header section ends at the first blank line (CRLF or LF).
  const headerEnd = rawMime.search(/\r?\n\r?\n/);
  const headerSection = headerEnd >= 0 ? rawMime.slice(0, headerEnd) : rawMime;
  // Unfold continuations.
  const unfolded = headerSection.replace(/\r?\n[ \t]+/g, " ");
  for (const line of unfolded.split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx < 0) continue;
    if (line.slice(0, idx).trim().toLowerCase() === lcName) {
      return line.slice(idx + 1).trim();
    }
  }
  return undefined;
}

// Strip MIME headers, return a best-effort body. If multipart we just
// return the original — the SDK normally yields a single text/html part.
export function bodyFromMime(rawMime: string): string {
  const headerEnd = rawMime.search(/\r?\n\r?\n/);
  return headerEnd >= 0 ? rawMime.slice(headerEnd).replace(/^\r?\n\r?\n/, "") : rawMime;
}

// Returns true if the message looks like a multipart/* body.
export function isMultipart(rawMime: string): boolean {
  const ct = readMimeHeader(rawMime, "Content-Type") ?? "";
  return /^multipart\//i.test(ct);
}

export interface ParsedMessage {
  htmlBody: string | null;
  plainBody: string | null;
  attachments: ParsedAttachment[];
}

export interface ParsedAttachment {
  name: string;
  type: string;
  data: Uint8Array;
}

// Pulls the user-facing body and any attachments out of a decrypted MIME
// blob. Handles a single text part or a multipart/* envelope. Decodes
// base64 and quoted-printable transfer encodings; other encodings (7bit,
// 8bit, binary) are passed through as-is.
export function parseDecryptedMime(rawMime: string): ParsedMessage {
  const result: ParsedMessage = { htmlBody: null, plainBody: null, attachments: [] };
  collectParts(rawMime, result);
  return result;
}

interface RawPart {
  headers: Record<string, string>;
  body: string;
}

function splitHeadersAndBody(raw: string): RawPart {
  const headerEnd = raw.search(/\r?\n\r?\n/);
  if (headerEnd < 0) return { headers: {}, body: raw };
  const headerSection = raw.slice(0, headerEnd);
  const body = raw.slice(headerEnd).replace(/^\r?\n\r?\n/, "");
  const unfolded = headerSection.replace(/\r?\n[ \t]+/g, " ");
  const headers: Record<string, string> = {};
  for (const line of unfolded.split(/\r?\n/)) {
    const idx = line.indexOf(":");
    if (idx < 0) continue;
    const name = line.slice(0, idx).trim().toLowerCase();
    headers[name] = line.slice(idx + 1).trim();
  }
  return { headers, body };
}

function collectParts(raw: string, out: ParsedMessage): void {
  const part = splitHeadersAndBody(raw);
  const ct = part.headers["content-type"] ?? "text/plain";
  const ctMain = ct.split(";")[0].trim().toLowerCase();

  if (ctMain.startsWith("multipart/")) {
    const boundary = paramFromHeader(ct, "boundary");
    if (!boundary) return;
    for (const child of splitByBoundary(part.body, boundary)) {
      collectParts(child, out);
    }
    return;
  }

  const cd = part.headers["content-disposition"] ?? "";
  const cte = (part.headers["content-transfer-encoding"] ?? "7bit").toLowerCase();
  const filename = paramFromHeader(cd, "filename") ?? paramFromHeader(ct, "name");
  const isAttachment =
    /attachment/i.test(cd) ||
    (filename != null && !ctMain.startsWith("text/"));

  if (isAttachment) {
    out.attachments.push({
      name: filename ?? "attachment",
      type: ctMain,
      data: decodeToBytes(part.body, cte),
    });
    return;
  }

  const text = decodeToString(part.body, cte);
  if (ctMain === "text/html" && out.htmlBody == null) {
    out.htmlBody = text;
  } else if (ctMain === "text/plain" && out.plainBody == null) {
    out.plainBody = text;
  } else if (out.htmlBody == null && out.plainBody == null) {
    // Unknown text-ish content; treat as plain.
    out.plainBody = text;
  }
}

function paramFromHeader(value: string, name: string): string | null {
  const re = new RegExp(`(?:^|;)\\s*${name}\\s*=\\s*"?([^";]+)"?`, "i");
  const m = value.match(re);
  return m ? m[1].trim() : null;
}

function splitByBoundary(body: string, boundary: string): string[] {
  const escaped = boundary.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  // Boundaries appear on their own line: --boundary, with --boundary--
  // marking the close. We split on either form, then drop the preamble
  // (before the first boundary) and epilogue (after the close).
  const re = new RegExp(`(?:^|\\r?\\n)--${escaped}(?:--)?[^\\n]*\\r?\\n?`, "g");
  const pieces = body.split(re);
  // First element is the preamble (before first boundary), discard.
  // The last may be the epilogue or the after-close section, also drop empties.
  const parts = pieces.slice(1).filter((p) => p.length > 0);
  return parts;
}

function decodeToBytes(body: string, encoding: string): Uint8Array {
  if (encoding === "base64") {
    const cleaned = body.replace(/\s/g, "");
    try {
      const bin = atob(cleaned);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return bytes;
    } catch {
      return new Uint8Array();
    }
  }
  if (encoding === "quoted-printable") {
    return new TextEncoder().encode(decodeQuotedPrintable(body));
  }
  return new TextEncoder().encode(body);
}

function decodeToString(body: string, encoding: string): string {
  if (encoding === "base64") {
    const cleaned = body.replace(/\s/g, "");
    try {
      const bin = atob(cleaned);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      return new TextDecoder("utf-8").decode(bytes);
    } catch {
      return body;
    }
  }
  if (encoding === "quoted-printable") {
    return decodeQuotedPrintable(body);
  }
  return body;
}

function decodeQuotedPrintable(s: string): string {
  return s
    .replace(/=\r?\n/g, "")
    .replace(/=([0-9A-F]{2})/gi, (_m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

export const POSTGUARD_ENCRYPTED_FILENAME = "postguard.encrypted";
export const POSTGUARD_HEADER = "x-postguard";
export const POSTGUARD_HEADER_VALUE = "0.1.0";
