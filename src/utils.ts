import type { AttributeCon, AttributeRequest } from "./types";

export const PKG_URL =
  process.env.PKG_URL || "https://postguard-main.cs.ru.nl/pkg";
export const EMAIL_ATTRIBUTE_TYPE = "pbdf.sidn-pbdf.email.email";
export const POSTGUARD_SUBJECT = "PostGuard Encrypted Email";
export const PG_ATTACHMENT_NAME = "postguard.encrypted";

export const PG_ARMOR_BEGIN = "-----BEGIN POSTGUARD MESSAGE-----";
export const PG_ARMOR_END = "-----END POSTGUARD MESSAGE-----";
export const PG_ARMOR_DIV_ID = "postguard-armor";
export const POSTGUARD_WEBSITE_URL =
  process.env.POSTGUARD_WEBSITE_URL || "https://postguard.eu";
export const PG_MAX_URL_FRAGMENT_SIZE = 100_000;

const EXT_VERSION = "0.2.0";

export const PG_CLIENT_HEADER: Record<string, string> = {
  "X-PostGuard-Client-Version": `Outlook,web,pg4outlook,${EXT_VERSION}`,
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
    ? `<div style="background:#F4F3F4;text-align:left;padding:20px;margin:40px 0;font-size:12px;">${htmlContent}</div>`
    : "";

  return `<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width,initial-scale=1">
    <meta name="x-apple-disable-message-reformatting">
    <title></title>
</head>
<body style="background:#F4F3F4;font-family:Karla,sans-serif;line-height:25px">
    <div style="width:100%;max-width:600px;margin-left:auto;margin-right:auto;text-align:center;">
        <div style="margin:50px 0 10px 0">
            <svg width="84" height="46" viewBox="0 0 84 46" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M44.8262 24.7016V10.9611H49.7197C50.8164 10.9611 51.6825 11.161 52.318 11.5598C52.9535 11.9587 53.4061 12.4728 53.6757 13.1004C53.9453 13.729 54.0797 14.3806 54.0797 15.0543C54.0797 15.5243 54.0006 15.999 53.8453 16.4776C53.689 16.9543 53.4385 17.3955 53.0916 17.7972C52.7448 18.1989 52.2961 18.5238 51.7444 18.7688C51.1927 19.0139 50.5191 19.1369 49.7207 19.1369H46.1887V24.7016H44.8271H44.8262Z" fill="#022E3D"/>
                <path d="M35.5799 24.2392C35.638 24.0749 35.6847 23.899 35.7428 23.7232C37.8475 16.4026 38.5107 10.0797 38.5107 10.0797L38.503 10.0768C34.4032 8.12104 23.448 2.45261 19.5434 0.587158C14.4955 3.22051 5.707 7.41753 0.578125 10.0778H0.580031C0.580031 10.0778 3.9491 42.1489 19.5434 45.4118H19.5415C27.2058 43.8058 31.904 35.2541 34.7186 26.9725" fill="#0071EB"/>
                <path d="M19.542 46C19.4868 46 19.4315 45.9923 19.381 45.9769C3.59134 42.6141 0.142239 11.4647 0.00313164 10.1393C-0.0216409 9.89907 0.102222 9.66649 0.315647 9.55597C2.82625 8.25275 6.27155 6.55356 9.60346 4.90916C13.0935 3.18883 16.7018 1.40892 19.2781 0.0653362C19.4372 -0.0182774 19.6306 -0.0221217 19.7936 0.0557255C21.7163 0.974515 25.2578 2.76981 29.0061 4.66985C32.7744 6.57951 36.6713 8.55356 38.7522 9.54636C39.0238 9.62516 39.119 9.85678 39.0895 10.1413C39.0828 10.2057 38.3968 16.5997 36.3026 23.8866L36.2273 24.1259C36.1959 24.2326 36.1625 24.3383 36.1273 24.4382C36.0167 24.7429 35.6871 24.8976 35.3822 24.791C35.0802 24.6814 34.9229 24.344 35.0325 24.0394C35.062 23.9567 35.0878 23.8712 35.1145 23.7837L35.1926 23.5339C36.9514 17.4137 37.7013 11.9231 37.8852 10.4209C35.6356 9.34068 32.0045 7.50118 28.4849 5.71743C24.9053 3.90291 21.5153 2.18547 19.5544 1.23881C16.9828 2.57471 13.4813 4.3008 10.0922 5.97211C6.92897 7.53194 3.66376 9.1427 1.2027 10.4113C1.66861 14.1922 5.53409 41.7318 19.5449 44.8111C25.6389 43.4675 30.5582 37.4012 34.1693 26.7804C34.2722 26.4738 34.6028 26.3094 34.9077 26.4161C35.2116 26.5199 35.3736 26.8553 35.2688 27.161C31.491 38.272 26.2402 44.6064 19.6592 45.9856C19.6192 45.9952 19.5801 45.9981 19.5411 45.9981L19.542 46Z" fill="#022E3D"/>
            </svg>
        </div>
        <div style="background:#fff;padding:80px;">
            <p style="color:#00A2D5;font-size:21px;margin-top:0;margin-bottom:5px;">
              ${sender}
            </p>
            <p style="font-size:21px;margin-top:5px;">
              sent you an encrypted email
            </p>
            ${htmlContentSection}
            <a href="${decryptUrl}" style="display:inline-block;font-weight:600;margin:20px 0;max-width:300px;width:100%;background:#00A2D5;border:none;border-radius:31px;color:#fff;padding:.7em 0;text-decoration:none;font-size:12px;">
              Decrypt this email
            </a>
            <div style="text-align:left;padding-top:40px;border-top:2px solid #F4F3F4">
                <p>Decrypt link</p>
                <a style="color:#00A2D5;font-size:12px;font-weight:700;line-height:14px;" href="${decryptUrl}">
                  ${decryptUrl}
                </a>
            </div>
        </div>
    </div>
    <div id="${PG_ARMOR_DIV_ID}" style="display:none;font-size:0;max-height:0;overflow:hidden;mso-hide:all">${armorBlock}</div>
</body>
</html>`;
}
