import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-microsoft-storage-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { db } = await import("../lib/db.ts");

const tenantId = "11111111-1111-4111-8111-111111111111";
const clientId = "22222222-2222-4222-8222-222222222222";

await storage.saveConfig({
  microsoftTenantId: tenantId,
  microsoftClientId: clientId,
  microsoftClientSecret: "synthetic-secret-one",
  microsoftTestUserEmail: "pilot@example.edu",
}, "test:setup");

test("admin-consent state is tenant-bound, one-time, and credential-version-bound", async () => {
  const connection = await storage.getStoredMicrosoftConnection();
  const state = await storage.createMicrosoftConsentState(connection, "test:administrator");
  const consent = await storage.consumeMicrosoftConsentState(state, tenantId.toUpperCase());

  assert.deepEqual(consent, {
    actor: "test:administrator",
    tenantId,
    clientId,
    credentialVersion: connection.credentialVersion,
  });
  await assert.rejects(
    storage.consumeMicrosoftConsentState(state, tenantId),
    /expired or was already used/,
  );
});

test("a late connection probe cannot verify credentials rotated while it was running", async () => {
  const probedConnection = await storage.getStoredMicrosoftConnection();
  await storage.recordMicrosoftAdminConsent("test:administrator", probedConnection);
  assert.notEqual((await storage.getConfig(false)).microsoftConsentGrantedAt, "");

  await storage.saveConfig({ microsoftClientSecret: "synthetic-secret-two" }, "test:administrator");
  const rotated = await storage.getConfig(false);
  assert.equal(rotated.microsoftConsentGrantedAt, "");

  await assert.rejects(
    storage.recordMicrosoftAdminConsent("test:administrator", probedConnection),
    /credentials changed during verification/,
  );
  assert.equal((await storage.getConfig(false)).microsoftConsentGrantedAt, "");

  const replacement = await storage.getStoredMicrosoftConnection();
  assert.equal(replacement.credentialVersion, probedConnection.credentialVersion + 1);
  await storage.recordMicrosoftAdminConsent("test:administrator", replacement);
  assert.notEqual((await storage.getConfig(false)).microsoftConsentGrantedAt, "");
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
