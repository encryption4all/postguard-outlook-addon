// Token acquisition for Microsoft Graph using Office SSO.
//
// Strategy:
//   1. Try OfficeRuntime.auth.getAccessToken({ allowSignInPrompt, forMSGraphAccess: true })
//      — succeeds when the addin is configured for Nested App Authentication
//      and the tenant admin has consented. This returns a Graph-scoped JWT
//      directly (no on-behalf-of exchange required).
//   2. Otherwise reject — callers should handle the unavailable-Graph case
//      by skipping Graph-only features (sent copy, in-place replacement).
//
// To enable: register an Azure AD app, add Mail.ReadWrite + User.Read scopes,
// add the <WebApplicationInfo> block to manifest.xml and (re)deploy.

let cached: { token: string; expiresAt: number } | null = null;

export async function getGraphToken(allowPrompt = false): Promise<string | null> {
  const now = Date.now();
  if (cached && cached.expiresAt > now + 60_000) return cached.token;

  try {
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: allowPrompt,
      forMSGraphAccess: true,
      allowConsentPrompt: allowPrompt,
    });
    if (!token) return null;

    // JWT exp is in seconds since epoch. Best-effort decode; on failure,
    // assume 50-minute lifetime (Graph tokens are typically 60-90 min).
    const expSec = parseJwtExp(token);
    cached = {
      token,
      expiresAt: expSec ? expSec * 1000 : now + 50 * 60_000,
    };
    return token;
  } catch (_e) {
    return null;
  }
}

function parseJwtExp(jwt: string): number | null {
  try {
    const parts = jwt.split(".");
    if (parts.length !== 3) return null;
    const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
    return typeof payload.exp === "number" ? payload.exp : null;
  } catch (_e) {
    return null;
  }
}

export function clearTokenCache(): void {
  cached = null;
}
