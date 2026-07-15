// Tests for the sender-trust reconciliation helper
// (src/lib/verified-sender.ts).
//
// The MIME `From` header travels inside the encrypted payload and is fully
// sender-controlled. Only the PostGuard-signature sender is trustworthy, so
// the read view must present that one as authoritative and flag a header that
// claims a different address. These tests pin that reconciliation.

import { test } from "node:test";
import assert from "node:assert/strict";
import { reconcileSender, extractAddress, senderMetaLine } from "../src/lib/verified-sender.ts";

// Minimal stand-in for the i18n lookup: echoes the key so assertions can see
// which strings the meta line is built from.
const echo = (k: string): string => k;

test("extractAddress pulls the address out of a display-name header", () => {
  assert.equal(extractAddress("Alice <alice@example.com>"), "alice@example.com");
  assert.equal(extractAddress("bob@example.com"), "bob@example.com");
  assert.equal(extractAddress("  MiXeD@Example.COM  "), "mixed@example.com");
  assert.equal(extractAddress("no address here"), null);
  assert.equal(extractAddress(""), null);
});

test("verified sender is authoritative when the header agrees", () => {
  const trust = reconcileSender("Alice <alice@example.com>", "alice@example.com");
  assert.equal(trust.verified, "alice@example.com");
  assert.equal(trust.mismatch, false);
});

test("a header naming a different address is flagged as a mismatch", () => {
  const trust = reconcileSender("Alice <alice@bank.example>", "attacker@evil.example");
  assert.equal(trust.verified, "attacker@evil.example");
  assert.equal(trust.mismatch, true);
});

test("mismatch is case-insensitive on the address", () => {
  const trust = reconcileSender("ALICE@EXAMPLE.COM", "alice@example.com");
  assert.equal(trust.mismatch, false);
});

test("no verified sender means nothing is treated as verified", () => {
  const trust = reconcileSender("alice@example.com", null);
  assert.equal(trust.verified, null);
  assert.equal(trust.mismatch, false);
});

test("meta line shows the verified sender, not the claimed header", () => {
  const line = senderMetaLine(reconcileSender("Mallory <alice@example.com>", "bob@example.com"), echo);
  assert.match(line.from, /metaFromVerified/);
  assert.match(line.from, /bob@example\.com/);
  assert.doesNotMatch(line.from, /alice@example\.com/);
});

test("meta line surfaces a warning with the claimed header on mismatch", () => {
  const line = senderMetaLine(
    reconcileSender("Alice <alice@bank.example>", "attacker@evil.example"),
    echo
  );
  assert.ok(line.warning);
  assert.match(line.warning!, /senderMismatchWarning/);
  assert.match(line.warning!, /alice@bank\.example/);
});

test("no warning when the verified and claimed addresses agree", () => {
  const line = senderMetaLine(reconcileSender("alice@example.com", "alice@example.com"), echo);
  assert.equal(line.warning, null);
});

test("an unverified sender is labelled unverified with no warning", () => {
  const line = senderMetaLine(reconcileSender("alice@example.com", null), echo);
  assert.match(line.from, /metaFrom\b/);
  assert.match(line.from, /senderUnverified/);
  assert.equal(line.warning, null);
});

test("an empty from with no verified sender yields no meta line", () => {
  const line = senderMetaLine(reconcileSender("", null), echo);
  assert.equal(line.from, "");
  assert.equal(line.warning, null);
});
