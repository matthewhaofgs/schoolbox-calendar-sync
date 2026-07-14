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

test("inactive users and ambiguous Schoolbox addresses cannot be matched", () => {
  const first = { id: 1, email: "first@example.edu", altEmail: "shared@example.edu", enabled: true };
  const second = { id: 2, email: "shared@example.edu", enabled: true };
  const disabled = { id: 3, email: "disabled@example.edu", enabled: false };
  const deleted = { id: 4, altEmail: "deleted@example.edu", enabled: true, isDeleted: true };
  const index = indexActiveSchoolboxUsersByEmail([first, second, disabled, deleted]);

  assert.equal(index.get("first@example.edu"), first);
  assert.equal(index.has("shared@example.edu"), false);
  assert.equal(index.has("disabled@example.edu"), false);
  assert.equal(index.has("deleted@example.edu"), false);
});

test("the same user's duplicate primary and alternate value remains unique", () => {
  const user = { id: 1, email: "same@example.edu", altEmail: "same@example.edu", enabled: true };
  assert.equal(indexActiveSchoolboxUsersByEmail([user]).get("same@example.edu"), user);
});
