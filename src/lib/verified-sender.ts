// Reconciles the two "sender" identities a decrypted PostGuard message
// carries so the read view can present a trustworthy one.
//
// The MIME `From` header lives *inside* the encrypted payload and is fully
// controlled by whoever built the message, so it must never be shown as if
// the add-in vouches for it. The only sender identity the add-in can trust is
// the one bound to the PostGuard signature (`FriendlySender.email`). This
// helper makes that verified identity authoritative and flags the case where
// the claimed header names a different address (a spoofing signal).

export interface SenderTrust {
  // The cryptographically verified sender address, or null when the message
  // carried no verified sender (e.g. unsigned) — then the claimed header is
  // all we have and is presented as unverified.
  verified: string | null;
  // The raw MIME `From` header value as claimed inside the payload.
  claimed: string;
  // True when a verified sender exists and the claimed header names a
  // different address than the verified one.
  mismatch: boolean;
}

// Pull a bare, lowercased email address out of a MIME `From` header value,
// which may be `Display Name <addr@host>` or a bare `addr@host`. Returns null
// when no address can be isolated.
export function extractAddress(from: string): string | null {
  const angle = from.match(/<([^>]*)>/);
  const raw = (angle ? angle[1] : from).trim().toLowerCase();
  return /^[^@\s]+@[^@\s]+$/.test(raw) ? raw : null;
}

export function reconcileSender(claimedFrom: string, verified: string | null): SenderTrust {
  const claimed = (claimedFrom ?? "").trim();
  const verifiedAddr = verified?.trim() || null;
  const claimedAddr = extractAddress(claimed);
  const mismatch =
    verifiedAddr != null && claimedAddr != null && claimedAddr !== verifiedAddr.toLowerCase();
  return { verified: verifiedAddr, claimed, mismatch };
}

// Builds the sender meta text for the read view from a reconciled trust
// result. `translate` is the i18n lookup (passed in so this module stays
// free of Office/DOM imports and remains unit-testable). Returns the "from"
// line and, on a spoofing mismatch, a separate warning line.
export function senderMetaLine(
  trust: SenderTrust,
  translate: (key: string) => string
): { from: string; warning: string | null } {
  if (trust.verified) {
    return {
      from: `${translate("metaFromVerified")}: ${trust.verified}`,
      warning: trust.mismatch ? `${translate("senderMismatchWarning")} ${trust.claimed}` : null,
    };
  }
  if (!trust.claimed) return { from: "", warning: null };
  return {
    from: `${translate("metaFrom")}: ${trust.claimed} (${translate("senderUnverified")})`,
    warning: null,
  };
}
