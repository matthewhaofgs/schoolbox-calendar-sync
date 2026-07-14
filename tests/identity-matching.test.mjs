import assert from "node:assert/strict";
import { test } from "node:test";

import { indexActiveSchoolboxUsersByEmail } from "../lib/sync.ts";

test("Schoolbox identities can match by a unique primary or alternate email", () => {
  const primary = { id: 1, email: " Primary@Example.edu ", enabled: true };
  const alternate = { id: 2, email: "old@example.edu", altEmail: "Alternate@Example.edu", enabled: true };
  const index = indexActiveSchoolboxUsersByEmail([primary, alternate]);

  assert.equal(index.get("primary@example.edu"), primary);
  assert.equal(index.get("alternate@example.edu"), alternate);
  assert.equal(index.get("old@example.edu"), alternate);
});

test("unique primary emails take precedence over another user's alternate email", () => {
  const alternate = { id: 1, email: "first@example.edu", altEmail: "shared@example.edu", enabled: true };
  const primary = { id: 2, email: "shared@example.edu", enabled: true };
  const index = indexActiveSchoolboxUsersByEmail([alternate, primary]);

  assert.equal(index.get("shared@example.edu"), primary);
});

test("inactive users and ambiguous addresses at the same level cannot be matched", () => {
  const first = { id: 1, email: "duplicate-primary@example.edu", altEmail: "duplicate-alt@example.edu", enabled: true };
  const second = { id: 2, email: "duplicate-primary@example.edu", enabled: true };
  const third = { id: 3, email: "third@example.edu", altEmail: "duplicate-alt@example.edu", enabled: true };
  const disabled = { id: 4, email: "disabled@example.edu", enabled: false };
  const deleted = { id: 5, altEmail: "deleted@example.edu", enabled: true, isDeleted: true };
  const index = indexActiveSchoolboxUsersByEmail([first, second, third, disabled, deleted]);

  assert.equal(index.has("duplicate-primary@example.edu"), false);
  assert.equal(index.has("duplicate-alt@example.edu"), false);
  assert.equal(index.has("disabled@example.edu"), false);
  assert.equal(index.has("deleted@example.edu"), false);
});

test("the same user's duplicate primary and alternate value remains unique", () => {
  const user = { id: 1, email: "same@example.edu", altEmail: "same@example.edu", enabled: true };
  assert.equal(indexActiveSchoolboxUsersByEmail([user]).get("same@example.edu"), user);
});
