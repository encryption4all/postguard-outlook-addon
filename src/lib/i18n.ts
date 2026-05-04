// Tiny localization helper. Office Add-ins do not provide an equivalent of
// browser.i18n.getMessage(), so we ship strings inline and look up by key.
// Locale resolution falls back to en if a key is missing.

type Bundle = Record<string, string>;

const en: Bundle = {
  appName: "PostGuard",
  appDescription: "End-to-end email encryption using identity-based encryption and Yivi",
  composeActionTitle: "PostGuard Encryption",

  encryptionEnabled: "PostGuard encryption is enabled",
  encryptionDisabled: "Click to enable PostGuard encryption",

  composeSwitchBarEnabled: "PostGuard encryption is on",
  composeSwitchBarDisabled: "PostGuard encryption is off. Sensitive content? Turn it on.",
  manageAccess: "Manage Access",
  sign: "Sign",
  encryptAndSend: "Encrypt & Send",
  reencryptAndSend: "Re-encrypt & Send",
  encrypting: "Encrypting…",

  composeBccWarning: "PostGuard does not support BCC. Either remove BCC or disable PostGuard.",
  composeNoRecipients: "Add at least one recipient before encrypting.",
  composeNoSenderEmail: "Could not determine the sender email address.",

  decryptButton: "Decrypt",
  decryptingButton: "Decrypting…",

  displayScriptDecryptBar: "This mail is encrypted using PostGuard.",
  displayScriptWasEncryptedBar: "This mail was originally encrypted using PostGuard.",

  displayMessageTitle: "You received a PostGuard encrypted email from",
  displayMessageHeading: "You need to prove who you are to decrypt and read this email.",
  displayMessageQrPrefix: "Scan the QR code with the Yivi app to disclose your e-mail address.",
  displayMessageTitleSign: "Sign the e-mail",
  displayMessageHeadingSign: "You need to prove who you are to sign this email.",

  displayMessageYiviHelpHeader: "What is the Yivi app?",
  displayMessageYiviHelpBody:
    "The Yivi app is a separate privacy-friendly authentication app (which is used also for other authentication purposes).",
  displayMessageYiviHelpLinkText: "More information about Yivi",
  displayMessageYiviHelpDownloadHeader: "Download the free Yivi app",

  policyEditorTitle: "PostGuard — Manage Access",
  policyEditorTitleSign: "PostGuard — Sign",
  policyEditorSave: "Save",
  policyEditorCancel: "Cancel",

  notificationHeaderBadgesLabel: "This message was sent by",
  notificationComposeBadgesLabel: "Recipients will know you as",

  decryptionFailed:
    "Decryption failed: the disclosed attributes did not match. Make sure you verify the correct email address in your Yivi app.",
  decryptionError: "Decryption failed. Please try again.",
  encryptionError: "Encryption failed. Please try again.",
  networkError: "Could not connect to PostGuard server. Check your network connection.",
  startupError:
    "PostGuard failed to initialize. Encryption and decryption will not work until the issue is resolved.",
  sentCopyError: "Failed to save the sent copy of your encrypted message.",
  recipientUnknown: "This message was not encrypted for the mail account it was received on.",

  "pbdf.sidn-pbdf.email.email": "Email address",
  "pbdf.sidn-pbdf.mobilenumber.mobilenumber": "Mobile number",
  "pbdf.gemeente.personalData.surname": "Surname",
  "pbdf.gemeente.personalData.dateofbirth": "Date of birth",
};

const bundles: Record<string, Bundle> = { en };

export function t(key: string, fallback?: string): string {
  const locale = (Office?.context?.displayLanguage ?? "en").slice(0, 2).toLowerCase();
  const bundle = bundles[locale] ?? bundles.en;
  return bundle[key] ?? bundles.en[key] ?? fallback ?? key;
}
