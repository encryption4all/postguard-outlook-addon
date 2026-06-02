// Internet-header keys and related per-message constants shared between
// the taskpane and the OnMessageSend launchevent runtime. The two
// runtimes are separate WebViews and cannot read each other's memory, so
// every header name they exchange has to live in one place — otherwise
// a name change in one file silently breaks the handshake.
//
// Custom internet header names must start with "x-" (Office.js requirement).

// Per-draft "encrypt this message on send" flag. This header is the
// single source of truth at send time — the OnMessageSend handler only
// runs the encryption flow when it reads "true" here, so the user's
// global Encryption setting only acts as the default that
// OnNewMessageCompose seeds onto each new draft.
export const HEADER_ENCRYPT_ON_SEND = "x-pg-encrypt-on-send";

// Comma-joined sorted list of lowercase To+Cc emails captured at encrypt
// time. The handler compares this against the message's current recipients
// to refuse sending an encrypted blob to anyone who wasn't in the policy.
export const HEADER_ENCRYPTED_RECIPIENTS = "x-pg-encrypted-recipients";

// PostGuard interop marker, written to outbound encrypted messages. The
// Thunderbird addon writes the same header (background.ts:485) and uses it
// as the OnMessageRead filter for the Outlook add-in. Detection on the
// receive side is still primarily attachment + body armor, but the header
// is a third independent signal that survives any HTML sanitation OWA
// applies during send.
export const HEADER_POSTGUARD = "x-postguard";
export const POSTGUARD_VERSION = "0.1.0";

// Notification key for the persistent in-message "PostGuard is on/off"
// banner. Shared between the taskpane (which updates it when the user
// flips the toggle) and OnNewMessageCompose (which paints the initial
// state on every new draft).
export const ENCRYPTION_STATUS_NOTIFICATION_KEY = "postguard-encryption-status";
