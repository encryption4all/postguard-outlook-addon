// Server URLs are injected at build time via webpack DefinePlugin.
// The PostGuard SDK handles all PKG and Cryptify communication internally.

export const PKG_URL: string = process.env.PKG_URL as string;
export const CRYPTIFY_URL: string = process.env.CRYPTIFY_URL as string;
export const POSTGUARD_WEBSITE_URL: string = process.env.POSTGUARD_WEBSITE_URL as string;
// Add-in's own public origin (e.g. https://addin.postguard.eu/). Used by the
// OnMessageSend launchevent runtime to construct the Yivi dialog URL —
// window.location is unreliable there on New Outlook for Mac.
export const ADDIN_PUBLIC_URL: string = process.env.ADDIN_PUBLIC_URL as string;
// Injected from package.json at build time so the fourth field of the
// client-version header tracks the deployed release automatically.
export const ADDIN_VERSION: string = process.env.ADDIN_VERSION as string;

export const CLIENT_NAME = "Outlook";
export const CLIENT_ID = "pg4ol";

export function clientHeaders(addinVersion: string): Record<string, string> {
  return {
    "X-PostGuard-Client-Version": `${CLIENT_NAME},1.0,${CLIENT_ID},${addinVersion}`,
    // Identifies this add-in in cryptify's per-channel upload metrics.
    // Required because cryptify's detect_channel checks the Origin
    // header before User-Agent, and the add-in is served from
    // addin.*.postguard.eu — without this header it would be
    // misclassified as `website` / `staging-website`.
    "X-Cryptify-Source": "outlook",
  };
}
