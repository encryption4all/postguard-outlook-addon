import type { AttributeCon, AttributeRequest } from "./types";

export const POSTGUARD_WEBSITE_URL =
  process.env.POSTGUARD_WEBSITE_URL || "https://postguard.staging.yivi.app";
export const PKG_URL =
  process.env.PKG_URL || `${POSTGUARD_WEBSITE_URL}/pkg`;
export const POSTGUARD_LOGO_URL =
  process.env.POSTGUARD_LOGO_URL || `${POSTGUARD_WEBSITE_URL}/pg_logo_no_text.png`;
export const EMAIL_ATTRIBUTE_TYPE = "pbdf.sidn-pbdf.email.email";
export const POSTGUARD_SUBJECT = "PostGuard Encrypted Email";
export const PG_ATTACHMENT_NAME = "postguard.encrypted";

export const PG_ARMOR_BEGIN = "-----BEGIN POSTGUARD MESSAGE-----";
export const PG_ARMOR_END = "-----END POSTGUARD MESSAGE-----";
export const PG_ARMOR_DIV_ID = "postguard-armor";
export const PG_MAX_URL_FRAGMENT_SIZE = 100_000;

const EXT_VERSION = "0.2.0";

export const PG_CLIENT_HEADER: Record<string, string> = {
  "X-PostGuard-Client-Version": `Outlook,web,pg4ol,${EXT_VERSION}`,
};

export function toEmail(identity: string): string {
  const regex = /^(.*)<(.*)>$/;
  const match = identity.match(regex);
  const email = match ? match[2] : identity;
  return email.toLowerCase();
}

export function generateBoundary(): string {
  const rand = crypto.getRandomValues(new Uint8Array(16));
  const hex = Array.from(rand)
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
  return hex;
}

export async function hashCon(con: AttributeCon): Promise<string> {
  const sorted = [...con].sort(
    (att1: AttributeRequest, att2: AttributeRequest) =>
      att1.t.localeCompare(att2.t) || (att1.v ?? "").localeCompare(att2.v ?? "")
  );
  return await hashString(JSON.stringify(sorted));
}

export async function hashString(message: string): Promise<string> {
  const msgArray = new TextEncoder().encode(message);
  const hashBuffer = await crypto.subtle.digest("SHA-256", msgArray);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, "0")).join("");
}

export function secondsTill4AM(): number {
  const now = Date.now();
  const nextMidnight = new Date(now).setHours(24, 0, 0, 0);
  const secondsTillMidnight = Math.round((nextMidnight - now) / 1000);
  const secondsTill4AM = secondsTillMidnight + 4 * 60 * 60;
  return secondsTill4AM % (24 * 60 * 60);
}

export function typeToImage(t: string): string {
  switch (t) {
    case "pbdf.sidn-pbdf.email.email":
      return "envelope";
    case "pbdf.sidn-pbdf.mobilenumber.mobilenumber":
      return "phone";
    case "pbdf.pbdf.surfnet-2.id":
      return "education";
    case "pbdf.nuts.agb.agbcode":
      return "health";
    case "pbdf.gemeente.personalData.dateofbirth":
      return "calendar";
    default:
      return "personal";
  }
}

// JWT cache using localStorage (Office add-ins can use localStorage)
export async function checkJwtCache(con: AttributeCon): Promise<string> {
  const hash = await hashCon(con);
  const cached = localStorage.getItem(`pg-jwt-${hash}`);
  if (!cached) throw new Error("not found in cache");
  const entry = JSON.parse(cached);
  if (Date.now() / 1000 > entry.exp) {
    localStorage.removeItem(`pg-jwt-${hash}`);
    throw new Error("jwt has expired");
  }
  return entry.jwt;
}

export async function invalidateJwtCache(con: AttributeCon): Promise<void> {
  const hash = await hashCon(con);
  localStorage.removeItem(`pg-jwt-${hash}`);
}

export async function storeJwtCache(con: AttributeCon, jwt: string): Promise<void> {
  const hash = await hashCon(con);
  const decoded = parseJwt(jwt);
  localStorage.setItem(`pg-jwt-${hash}`, JSON.stringify({ jwt, exp: decoded.exp }));
}

function parseJwt(token: string): { exp: number } {
  const base64Url = token.split(".")[1];
  const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
  const jsonPayload = decodeURIComponent(
    atob(base64)
      .split("")
      .map((c) => "%" + ("00" + c.charCodeAt(0).toString(16)).slice(-2))
      .join("")
  );
  return JSON.parse(jsonPayload);
}

// Retrieve master public key from PKG, with localStorage fallback.
export async function retrievePublicKey(): Promise<string> {
  const PK_KEY = "pg-pk";
  const storedPublicKey = localStorage.getItem(PK_KEY);

  try {
    const resp = await fetch(`${PKG_URL}/v2/parameters`, {
      headers: PG_CLIENT_HEADER,
    });
    const { publicKey } = await resp.json();
    if (storedPublicKey !== publicKey) {
      localStorage.setItem(PK_KEY, publicKey);
    }
    return publicKey;
  } catch (e) {
    console.log(`[PostGuard] Failed to retrieve public key from PKG, falling back to cache`);
    if (storedPublicKey) return storedPublicKey;
    throw new Error("no public key");
  }
}

// Retrieve master verification key from PKG.
export async function retrieveVerificationKey(): Promise<string> {
  const resp = await fetch(`${PKG_URL}/v2/sign/parameters`);
  const { publicKey } = await resp.json();
  return publicKey;
}

// Request User Seal Key using JWT.
export async function getUSK(jwt: string, ts: number): Promise<unknown> {
  const url = `${PKG_URL}/v2/irma/key/${ts?.toString()}`;
  const resp = await fetch(url, {
    headers: {
      Authorization: `Bearer ${jwt}`,
      ...PG_CLIENT_HEADER,
    },
  });
  const json = await resp.json();
  if (json.status !== "DONE" || json.proofStatus !== "VALID") {
    throw new Error("session not DONE and VALID");
  }
  return json.key;
}

// Request signing keys using JWT.
export async function getSigningKeys(
  jwt: string,
  keyRequest?: object
): Promise<{ pubSignKey: unknown; privSignKey?: unknown }> {
  const url = `${PKG_URL}/v2/irma/sign/key`;
  const resp = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${jwt}`,
      ...PG_CLIENT_HEADER,
      "content-type": "application/json",
    },
    body: JSON.stringify(keyRequest),
  });
  const json = await resp.json();
  if (json.status !== "DONE" || json.proofStatus !== "VALID") {
    throw new Error("session not DONE and VALID");
  }
  return { pubSignKey: json.pubSignKey, privSignKey: json.privSignKey };
}

// Clean up expired JWT cache entries.
export function cleanUpCache(): void {
  const now = Date.now() / 1000;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (key?.startsWith("pg-jwt-")) {
      try {
        const entry = JSON.parse(localStorage.getItem(key)!);
        if (now > entry.exp) {
          localStorage.removeItem(key);
        }
      } catch {
        // ignore malformed entries
      }
    }
  }
}

// ─── MIME parsing ──────────────────────────────────────────────────

export function parseMimeContent(mimeData: string): { subject: string; body: string; isHtml: boolean } {
  const subjectMatch = mimeData.match(/^Subject:\s*(.+)$/im);
  const subject = subjectMatch ? subjectMatch[1].trim() : "(no subject)";

  const headerEndIndex = mimeData.indexOf("\r\n\r\n");
  let body = headerEndIndex !== -1 ? mimeData.substring(headerEndIndex + 4) : mimeData;

  const contentTypeMatch = mimeData.match(
    /^Content-Type:\s*multipart\/mixed;\s*boundary="?([^"\r\n]+)"?/im
  );
  if (contentTypeMatch) {
    const boundary = contentTypeMatch[1];
    const parts = body.split(`--${boundary}`);
    if (parts.length > 1) {
      const firstPart = parts[1];
      const partHeaderEnd = firstPart.indexOf("\r\n\r\n");
      body = partHeaderEnd !== -1 ? firstPart.substring(partHeaderEnd + 4) : firstPart;
    }
  }

  const isHtml = /<html|<div|<p|<br/i.test(body);
  return { subject, body, isHtml };
}

// ─── Armor & body helpers ──────────────────────────────────────────

export function armorBase64(base64: string): string {
  const lines: string[] = [];
  for (let i = 0; i < base64.length; i += 76) {
    lines.push(base64.substring(i, i + 76));
  }
  return `${PG_ARMOR_BEGIN}\n${lines.join("\n")}\n${PG_ARMOR_END}`;
}

export function extractArmoredPayload(html: string): string | null {
  const regex = new RegExp(
    PG_ARMOR_BEGIN.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") +
      "\\s*([A-Za-z0-9+/=\\s]+?)\\s*" +
      PG_ARMOR_END.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
  );
  const match = html.match(regex);
  if (!match) return null;
  return match[1].replace(/\s/g, "");
}

export function toUrlSafeBase64(base64: string): string {
  return base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export function fromUrlSafeBase64(urlSafe: string): string {
  let base64 = urlSafe.replace(/-/g, "+").replace(/_/g, "/");
  const pad = base64.length % 4;
  if (pad === 2) base64 += "==";
  else if (pad === 3) base64 += "=";
  return base64;
}

export function buildEncryptedBody(
  sender: string,
  base64Encrypted: string,
  htmlContent: string = ""
): string {
  let decryptUrl: string;
  if (base64Encrypted.length <= PG_MAX_URL_FRAGMENT_SIZE) {
    const urlSafe = toUrlSafeBase64(base64Encrypted);
    decryptUrl = `${POSTGUARD_WEBSITE_URL}/decrypt#${urlSafe}`;
  } else {
    decryptUrl = `${POSTGUARD_WEBSITE_URL}/decrypt`;
  }

  const armorBlock = armorBase64(base64Encrypted);

  const htmlContentSection = htmlContent
    ? `<div style="text-align:left;padding:20px 24px;margin:30px 0;font-size:14px;background:#F2F8FD;color:#030E17;line-height:22px;">${htmlContent}</div>`
    : "";

  return `<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <meta name="x-apple-disable-message-reformatting">
    <title></title>
</head>
<body style="background:#F2F8FD;background-color:#F2F8FD;font-family:Overpass,sans-serif;line-height:25px;color:#030E17;margin:0;padding:0">
    <div style="background:#F2F8FD;background-color:#F2F8FD;padding:1em;">
        <div style="background:#F2F8FD;width:100%;max-width:600px;margin-left:auto;margin-right:auto;text-align:center;">
            <div style="margin:50px 0 20px 0">
                <img src="${POSTGUARD_LOGO_URL}" alt="PostGuard" width="82" height="82" style="display:block;margin:0 auto;" />
            </div>
            <div style="background:#FFFFFF;padding:60px 50px;border-radius:8px;text-align:center;">
                <p style="font-size:22px;font-weight:700;color:#030E17;margin:0 0 5px 0;line-height:30px;">
                ${sender} sent you an encrypted email
                </p>
                ${htmlContentSection}
                <a href="${decryptUrl}" style="display:inline-block;font-weight:600;margin:25px 0;max-width:350px;width:100%;background:#030E17;border:none;border-radius:6px;color:#ffffff;padding:14px 0;text-decoration:none;font-size:16px;">
                Decrypt this email
                </a>
                <div style="text-align:left;padding-top:30px;">
                    <p style="color:#5F7381;font-size:16px;font-weight:600;margin:0 0 4px 0;">Decrypt link</p>
                    <a style="color:#3095DE;font-size:13px;font-weight:400;line-height:18px;word-break:break-all;" href="${decryptUrl}">
                    ${decryptUrl}
                    </a>
                </div>
            </div>
            <div style="height:40px;"></div>
        </div>
    </div>
    <div id="${PG_ARMOR_DIV_ID}" style="display:none;font-size:0;max-height:0;overflow:hidden;mso-hide:all">${armorBlock}</div>
</body>
</html>`;
}
