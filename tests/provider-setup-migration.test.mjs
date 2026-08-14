import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, test } from "node:test";
import Database from "better-sqlite3";

const temporary = mkdtempSync(join(tmpdir(), "relay-provider-setup-migration-"));
const databasePath = join(temporary, "relay.sqlite");
process.env.DATABASE_PATH = databasePath;
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const { encryptSecret } = await import("../lib/security.ts");
const encryptedSchoolboxToken = await encryptSecret("synthetic-schoolbox-token");
const encryptedGoogleCredential = await encryptSecret(JSON.stringify({
  type: "service_account",
  client_email: "relay@example-project.iam.gserviceaccount.com",
  client_id: "123456789012345678901",
  private_key: "-----BEGIN PRIVATE KEY-----\nsynthetic-test-key\n-----END PRIVATE KEY-----\n",
}));

const legacy = new Database(databasePath);
const now = new Date().toISOString();
legacy.exec(`
  CREATE TABLE app_config (
    id INTEGER PRIMARY KEY DEFAULT 1,
    schoolbox_base_url TEXT,
    schoolbox_token_encrypted TEXT,
    google_service_account_encrypted TEXT,
    google_admin_email TEXT,
    google_customer TEXT NOT NULL DEFAULT 'my_customer',
    timezone TEXT NOT NULL DEFAULT 'Australia/Sydney',
    past_days INTEGER NOT NULL DEFAULT 30,
    future_days INTEGER NOT NULL DEFAULT 180,
    concurrency INTEGER NOT NULL DEFAULT 3,
    sync_interval_minutes INTEGER NOT NULL DEFAULT 360,
    enabled INTEGER NOT NULL DEFAULT 0,
    setup_completed INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL
  )
`);
legacy.prepare(`INSERT INTO app_config
  (id, schoolbox_base_url, schoolbox_token_encrypted,
   google_service_account_encrypted, google_admin_email,
   enabled, setup_completed, updated_at)
  VALUES (1, ?, ?, ?, ?, 1, 1, ?)`)
  .run(
    "https://schoolbox.example.edu",
    encryptedSchoolboxToken,
    encryptedGoogleCredential,
    "calendar-admin@example.edu",
    now,
  );
legacy.close();

const storage = await import("../lib/storage.ts");
const { db } = await import("../lib/db.ts");

test("legacy completed Google installations infer independent provider readiness", async () => {
  await storage.ensureSchema();
  const config = await storage.getConfig(false);

  assert.equal(config.schoolboxConfigured, true);
  assert.equal(config.schoolboxSetupCompleted, true);
  assert.equal(config.googleConfigured, true);
  assert.equal(config.googleSetupCompleted, true);
  assert.equal(config.googleEnabled, true);
  assert.equal(config.microsoftConfigured, false);
  assert.equal(config.microsoftSetupCompleted, false);
  assert.equal(config.microsoftEnabled, false);
  assert.equal(config.setupCompleted, true);
  assert.equal(config.enabled, true);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
