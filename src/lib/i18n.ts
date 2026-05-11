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
  composeEncryptionOnBanner: "PostGuard is on — this message will be encrypted on send.",
  composeEncryptionOffBanner: "PostGuard is off — this message will be sent unencrypted.",
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
  errorRetry: "Retry",
  dialogClose: "Close",
  decryptedAttachmentsHeading: "Attachments",
  removeRecipient: "Remove",
  loading: "Loading",
  metaFrom: "From",
  metaDate: "Date",
  readNoopMessage: "This message is not encrypted with PostGuard.",
  yiviCancel: "Cancel",

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

  settingsTitle: "PostGuard Settings",
  settingsEncryptionTitle: "Encryption",
  settingsEncryptionDefaultLabel: "Encrypt new messages by default",
  settingsEncryptionDefaultHelp:
    "Off by default. When on, every new message starts with PostGuard encryption enabled and shows a banner in the message until you turn it off in the compose taskpane.",
  settingsPrefillTitle: "Sender attributes",
  settingsPrefillHelp:
    "These attributes are always offered when you sign a message. Fill in a value to make the disclosure mandatory (the Yivi app must match it). Leave blank to keep the attribute optional — you decide in the Yivi app at signing time.",
  settingsAdvancedTitle: "Advanced",
  settingsAllowOptimisticDialogLabel: "Skip the “open a dialog” confirmation",
  settingsAllowOptimisticDialogHelp:
    "Off by default — Outlook shows a “PostGuard wants to open a dialog” prompt before each send. Turn this on only if you have already allowed pop-ups for this add-in in your browser’s site settings. Safari: Settings → Websites → Pop-ups → Allow for this site.",
  settingsBack: "Back",
  settingsSave: "Save",
  settingsSaved: "Saved.",
  settingsSaveError: "Could not save settings. Try again.",
  settingsOpen: "Settings",

  "pbdf.sidn-pbdf.email.email": "Email address",
  "pbdf.sidn-pbdf.mobilenumber.mobilenumber": "Mobile number",
  "pbdf.gemeente.personalData.surname": "Surname",
  "pbdf.gemeente.personalData.fullname": "Name",
  "pbdf.gemeente.personalData.dateofbirth": "Date of birth",
};

const nl: Bundle = {
  appName: "PostGuard",
  appDescription: "End-to-end e-mailversleuteling met identiteitsgebaseerde encryptie en Yivi",
  composeActionTitle: "PostGuard-versleuteling",

  encryptionEnabled: "PostGuard-versleuteling staat aan",
  encryptionDisabled: "Klik om PostGuard-versleuteling in te schakelen",

  composeSwitchBarEnabled: "PostGuard-versleuteling staat aan",
  composeSwitchBarDisabled: "PostGuard-versleuteling staat uit. Gevoelige inhoud? Schakel het in.",
  composeEncryptionOnBanner: "PostGuard staat aan — dit bericht wordt bij verzenden versleuteld.",
  composeEncryptionOffBanner: "PostGuard staat uit — dit bericht wordt onversleuteld verzonden.",
  manageAccess: "Toegang beheren",
  sign: "Ondertekenen",
  encryptAndSend: "Versleutelen en verzenden",
  reencryptAndSend: "Opnieuw versleutelen en verzenden",
  encrypting: "Bezig met versleutelen…",

  composeBccWarning: "PostGuard ondersteunt geen BCC. Verwijder BCC of schakel PostGuard uit.",
  composeNoRecipients: "Voeg ten minste één ontvanger toe voordat je versleutelt.",
  composeNoSenderEmail: "Het e-mailadres van de afzender kon niet worden bepaald.",

  decryptButton: "Ontsleutelen",
  decryptingButton: "Bezig met ontsleutelen…",

  displayScriptDecryptBar: "Deze e-mail is versleuteld met PostGuard.",
  displayScriptWasEncryptedBar: "Deze e-mail was oorspronkelijk versleuteld met PostGuard.",

  displayMessageTitle: "Je hebt een met PostGuard versleutelde e-mail ontvangen van",
  displayMessageHeading: "Je moet bewijzen wie je bent om deze e-mail te ontsleutelen en te lezen.",
  displayMessageQrPrefix: "Scan de QR-code met de Yivi-app om je e-mailadres te onthullen.",
  displayMessageTitleSign: "Onderteken de e-mail",
  displayMessageHeadingSign: "Je moet bewijzen wie je bent om deze e-mail te ondertekenen.",

  displayMessageYiviHelpHeader: "Wat is de Yivi-app?",
  displayMessageYiviHelpBody:
    "De Yivi-app is een aparte privacyvriendelijke authenticatie-app (die ook voor andere authenticatiedoeleinden wordt gebruikt).",
  displayMessageYiviHelpLinkText: "Meer informatie over Yivi",
  displayMessageYiviHelpDownloadHeader: "Download de gratis Yivi-app",

  policyEditorTitle: "PostGuard — Toegang beheren",
  policyEditorTitleSign: "PostGuard — Ondertekenen",
  policyEditorSave: "Opslaan",
  policyEditorCancel: "Annuleren",
  errorRetry: "Opnieuw proberen",
  dialogClose: "Sluiten",
  decryptedAttachmentsHeading: "Bijlagen",
  removeRecipient: "Verwijderen",
  loading: "Bezig met laden",
  metaFrom: "Van",
  metaDate: "Datum",
  readNoopMessage: "Dit bericht is niet versleuteld met PostGuard.",
  yiviCancel: "Annuleren",

  notificationHeaderBadgesLabel: "Dit bericht is verzonden door",
  notificationComposeBadgesLabel: "Ontvangers herkennen je als",

  decryptionFailed:
    "Ontsleutelen mislukt: de onthulde attributen kwamen niet overeen. Controleer of je in de Yivi-app het juiste e-mailadres gebruikt.",
  decryptionError: "Ontsleutelen mislukt. Probeer het opnieuw.",
  encryptionError: "Versleutelen mislukt. Probeer het opnieuw.",
  networkError:
    "Kan geen verbinding maken met de PostGuard-server. Controleer je netwerkverbinding.",
  startupError:
    "PostGuard kon niet worden geïnitialiseerd. Versleutelen en ontsleutelen werken niet totdat dit is opgelost.",
  sentCopyError: "Kon de verzonden kopie van je versleutelde bericht niet opslaan.",
  recipientUnknown:
    "Dit bericht is niet versleuteld voor het e-mailaccount waarop het is ontvangen.",

  settingsTitle: "PostGuard-instellingen",
  settingsEncryptionTitle: "Versleuteling",
  settingsEncryptionDefaultLabel: "Nieuwe berichten standaard versleutelen",
  settingsEncryptionDefaultHelp:
    "Standaard uit. Wanneer aan, start elk nieuw bericht met PostGuard-versleuteling ingeschakeld en wordt een banner in het bericht getoond totdat je het uitschakelt in het opstel-taakvenster.",
  settingsPrefillTitle: "Afzenderattributen",
  settingsPrefillHelp:
    "Deze attributen worden altijd aangeboden wanneer je een bericht ondertekent. Vul een waarde in om de onthulling verplicht te maken (de Yivi-app moet exact deze waarde tonen). Laat leeg om het attribuut optioneel te houden — dan beslis je in de Yivi-app of je het onthult.",
  settingsAdvancedTitle: "Geavanceerd",
  settingsAllowOptimisticDialogLabel: "Sla de bevestiging “dialoogvenster openen” over",
  settingsAllowOptimisticDialogHelp:
    "Standaard uit — Outlook toont vóór elke verzending een melding “PostGuard wil een dialoogvenster openen”. Schakel dit alleen in als je pop-ups voor deze add-in al hebt toegestaan in de site-instellingen van je browser. Safari: Instellingen → Websites → Pop-upvensters → Toestaan voor deze site.",
  settingsBack: "Terug",
  settingsSave: "Opslaan",
  settingsSaved: "Opgeslagen.",
  settingsSaveError: "Instellingen konden niet worden opgeslagen. Probeer het opnieuw.",
  settingsOpen: "Instellingen",

  "pbdf.sidn-pbdf.email.email": "E-mailadres",
  "pbdf.sidn-pbdf.mobilenumber.mobilenumber": "Mobiel nummer",
  "pbdf.gemeente.personalData.surname": "Achternaam",
  "pbdf.gemeente.personalData.fullname": "Naam",
  "pbdf.gemeente.personalData.dateofbirth": "Geboortedatum",
};

const bundles: Record<string, Bundle> = { en, nl };

export function t(key: string, fallback?: string): string {
  const locale = (Office?.context?.displayLanguage ?? "en").slice(0, 2).toLowerCase();
  const bundle = bundles[locale] ?? bundles.en;
  return bundle[key] ?? bundles.en[key] ?? fallback ?? key;
}
