import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, test } from "node:test";

const temporary = mkdtempSync(join(tmpdir(), "relay-provider-setup-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { runFullSync } = await import("../lib/sync.ts");
const { db } = await import("../lib/db.ts");

const serviceAccountJson = JSON.stringify({
  type: "service_account",
  client_email: "relay@example-project.iam.gserviceaccount.com",
  client_id: "123456789012345678901",
  private_key: "-----BEGIN PRIVATE KEY-----\nsynthetic-test-key\n-----END PRIVATE KEY-----\n",
});
const tenantId = "11111111-1111-4111-8111-111111111111";
const clientId = "22222222-2222-4222-8222-222222222222";

test("provider setup can be verified and completed independently", async () => {
  let config = await storage.getConfig(false);
  assert.equal(config.schoolboxSetupCompleted, false);
  assert.equal(config.googleSetupCompleted, false);
  assert.equal(config.microsoftSetupCompleted, false);

  await storage.saveConfig({
    schoolboxBaseUrl: "https://schoolbox.example.edu",
    schoolboxToken: "synthetic-schoolbox-token",
    schoolboxSetupCompleted: false,
  }, "test:administrator");
  const schoolbox = await storage.getStoredSchoolboxConnection();
  await assert.rejects(
    storage.saveConfig({ schoolboxSetupCompleted: true }, "test:administrator"),
    /Test the saved Schoolbox connection/,
  );
  await storage.recordConnectionVerified("schoolbox", "test:administrator", schoolbox.credentialVersion);
  config = await storage.saveConfig({ schoolboxSetupCompleted: true }, "test:administrator");
  assert.equal(config.schoolboxSetupCompleted, true);
  assert.equal(config.setupCompleted, false, "a source connection alone is not operational setup");

  await storage.saveConfig({
    googleServiceAccountJson: serviceAccountJson,
    googleAdminEmail: "admin@example.edu",
    googleCustomer: "C01234567",
    googleSetupCompleted: false,
  }, "test:administrator");
  const google = await storage.getStoredGoogleConnection();
  await assert.rejects(
    storage.saveConfig({ googleSetupCompleted: true }, "test:administrator"),
    /Test the saved Google Workspace connection/,
  );
  await storage.recordConnectionVerified("google", "test:administrator", google.credentialVersion);
  config = await storage.saveConfig({ googleSetupCompleted: true, googleEnabled: true }, "test:administrator");
  assert.equal(config.googleSetupCompleted, true);
  assert.equal(config.microsoftSetupCompleted, false);
  assert.equal(config.setupCompleted, true);

  await storage.saveConfig({
    microsoftTenantId: tenantId,
    microsoftClientId: clientId,
    microsoftClientSecret: "synthetic-microsoft-secret",
    microsoftTestUserEmail: "pilot@example.edu",
    microsoftSetupCompleted: false,
  }, "test:administrator");
  const microsoft = await storage.getStoredMicrosoftConnection();
  await storage.recordMicrosoftAdminConsent("test:administrator", microsoft);
  config = await storage.saveConfig({ microsoftSetupCompleted: true, microsoftEnabled: true }, "test:administrator");
  assert.equal(config.googleSetupCompleted, true);
  assert.equal(config.googleEnabled, true);
  assert.equal(config.microsoftSetupCompleted, true);
  assert.equal(config.microsoftEnabled, true);
});

test("connection changes invalidate and pause only the changed provider", async () => {
  let config = await storage.saveConfig({ enabled: true }, "test:administrator");
  assert.equal(config.enabled, true);

  config = await storage.saveConfig({ googleCustomer: "C07654321", googleEnabled: true }, "test:administrator");
  assert.equal(config.googleSetupCompleted, false);
  assert.equal(config.googleEnabled, false);
  assert.equal(config.microsoftSetupCompleted, true);
  assert.equal(config.microsoftEnabled, true);
  assert.equal(config.enabled, true, "provider rotation must preserve the scheduler preference");
  assert.equal(config.setupCompleted, true, "the other completed target keeps Relay operational");

  const google = await storage.getStoredGoogleConnection();
  await storage.recordConnectionVerified("google", "test:administrator", google.credentialVersion);
  config = await storage.saveConfig({ googleSetupCompleted: true, googleEnabled: true }, "test:administrator");
  assert.equal(config.googleSetupCompleted, true);

  config = await storage.saveConfig({ microsoftTestUserEmail: "replacement@example.edu" }, "test:administrator");
  assert.equal(config.microsoftConsentGrantedAt.length > 0, true, "a mailbox-only change does not revoke tenant consent");
  assert.equal(config.microsoftSetupCompleted, false);
  assert.equal(config.microsoftEnabled, false);
  assert.equal(config.googleSetupCompleted, true);
  assert.equal(config.googleEnabled, true);
  assert.equal(config.enabled, true);

  config = await storage.saveConfig({ futureDays: 120 }, "test:administrator");
  assert.equal(config.googleSetupCompleted, true, "policy and schedule edits do not invalidate provider verification");
  config = await storage.saveConfig({ googleEnabled: false }, "test:administrator");
  assert.equal(config.googleSetupCompleted, true, "disabling a target preserves its completed setup");
});

test("a late diagnostic cannot verify a connection changed during its probe", async () => {
  const probed = await storage.getStoredGoogleConnection();
  await storage.saveConfig({ googleAdminEmail: "replacement-admin@example.edu" }, "test:administrator");
  await assert.rejects(
    storage.recordConnectionVerified("google", "test:administrator", probed.credentialVersion),
    /settings changed during verification/,
  );
  assert.equal((await storage.getConfig(false)).googleSetupCompleted, false);
});

test("an incomplete provider cannot be activated or manually synchronized", async () => {
  let config = await storage.saveConfig({ googleEnabled: true }, "test:administrator");
  assert.equal(config.googleSetupCompleted, false);
  assert.equal(config.googleEnabled, false, "activation requests are ineffective until setup is verified");
  assert.equal(config.enabled, true, "the scheduler preference is preserved while no target is ready");

  await assert.rejects(
    runFullSync("manual", "test:administrator", {}, {}, ["google"]),
    /requested calendar targets are disabled/,
  );

  // Defend against legacy or externally corrupted state where an incomplete
  // target's raw enabled bit is true despite the storage invariant.
  db().prepare("UPDATE app_config SET google_sync_enabled = 1, google_setup_completed = 0 WHERE id = 1").run();
  await assert.rejects(
    runFullSync("manual", "test:administrator", {}, {}, ["google"]),
    /have not completed setup and verification/,
  );
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
