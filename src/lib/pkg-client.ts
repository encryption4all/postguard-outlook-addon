// Server URLs are injected at build time via webpack DefinePlugin.
// The PostGuard SDK handles all PKG and Cryptify communication internally.

export const PKG_URL: string = process.env.PKG_URL as string;
export const CRYPTIFY_URL: string = process.env.CRYPTIFY_URL as string;
export const POSTGUARD_WEBSITE_URL: string = process.env.POSTGUARD_WEBSITE_URL as string;

export const CLIENT_NAME = "Outlook";
export const CLIENT_ID = "pg4ol";

export function clientHeaders(addinVersion: string): Record<string, string> {
  return {
    "X-PostGuard-Client-Version": `${CLIENT_NAME},1.0,${CLIENT_ID},${addinVersion}`,
  };
}
