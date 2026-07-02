// Tests for the OData $filter escaping in the Graph client
// (src/lib/graph-client.ts).
//
// Run with Node's built-in test runner (no extra dependencies):
//   npm test
//
// These pin the OData-injection fix (advisory GHSA-xwmw-r3pq-2rm5, issue #116):
// a folder displayName containing a single quote must be escaped by doubling
// it so it cannot break out of the $filter string literal.

import { test } from "node:test";
import assert from "node:assert/strict";
import { escapeODataString } from "../src/lib/odata.ts";

// Reproduces the exact literal graph-client builds so the assertions describe
// the value Graph actually receives.
function buildFilterLiteral(displayName: string): string {
  return `displayName eq '${encodeURIComponent(escapeODataString(displayName))}'`;
}

test("happy path: a plain name is unchanged and stays inside the literal", () => {
  assert.equal(escapeODataString("PostGuard"), "PostGuard");
  assert.equal(buildFilterLiteral("PostGuard"), "displayName eq 'PostGuard'");
});

test("rejection path: a single quote is doubled, not left to close the literal", () => {
  // Without escaping this would be `... eq 'O'Brien'` — the second quote closes
  // the literal and `Brien'` becomes trailing OData the attacker controls.
  assert.equal(escapeODataString("O'Brien"), "O''Brien");
  // encodeURIComponent does NOT encode single quotes, so the doubled quote is
  // what actually protects the literal on the wire.
  assert.equal(buildFilterLiteral("O'Brien"), "displayName eq 'O''Brien'");
});

test("injection attempt cannot break out of the string literal", () => {
  // A crafted name trying to append its own OData clause.
  const malicious = "x' or startswith(displayName,'a";
  const escaped = escapeODataString(malicious);
  // Every original quote survives only as a doubled (escaped) quote; there is
  // no lone quote left that could terminate the literal early.
  assert.doesNotMatch(escaped.replace(/''/g, ""), /'/);
  assert.equal(escaped, "x'' or startswith(displayName,''a");
});

test("multiple quotes are each doubled", () => {
  assert.equal(escapeODataString("'''"), "''''''");
});
