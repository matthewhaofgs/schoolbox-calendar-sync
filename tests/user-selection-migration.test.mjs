import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import Database from "better-sqlite3";

const temporary = mkdtempSync(join(tmpdir(), "relay-user-selection-migration-"));
const databasePath = join(temporary, "relay.sqlite");
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
  );
  CREATE TABLE user_mappings (
    google_user_id TEXT PRIMARY KEY,
    google_email TEXT NOT NULL UNIQUE,
    schoolbox_user_id INTEGER,
    schoolbox_email TEXT,
    display_name TEXT,
    role TEXT,
    status TEXT NOT NULL DEFAULT 'unmatched',
    last_sync_at TEXT,
    last_error TEXT,
    event_count INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL
  );
  CREATE TABLE event_mappings (
    google_user_id TEXT NOT NULL,
    source_key TEXT NOT NULL,
    google_event_id TEXT NOT NULL,
    source_hash TEXT NOT NULL,
    source_start TEXT NOT NULL,
    source_end TEXT NOT NULL,
    last_seen_run_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (google_user_id, source_key)
  );
`);
legacy.prepare("INSERT INTO app_config (id, updated_at) VALUES (1, ?)").run(now);
legacy.prepare(`INSERT INTO user_mappings
  (google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name, role, status, last_sync_at, last_error, event_count, updated_at)
  VALUES ('existing-user', 'existing@example.edu', 42, 'existing@example.edu', 'Existing User', 'Staff', 'synced', ?, NULL, 3, ?)`)
  .run(now, now);
legacy.prepare(`INSERT INTO event_mappings
  (google_user_id, source_key, google_event_id, source_hash, source_start, source_end, last_seen_run_id, created_at, updated_at)
  VALUES ('existing-user', 'legacy-source', 'legacy-event', 'legacy-hash', ?, ?, 'legacy-run', ?, ?)`)
  .run(now, new Date(Date.now() + 60_000).toISOString(), now, now);
legacy.close();

process.env.DATABASE_PATH = databasePath;
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { db } = await import("../lib/db.ts");

await storage.ensureSchema();

test("upgrades preserve sync-all behavior for existing installations and mappings", async () => {
  const config = await storage.getConfig();
  const [mapping] = await storage.listUserMappings();

  assert.equal(config.syncNewUsersByDefault, true);
  assert.equal(config.syncPolicy.eventTypeMode, "all");
  assert.ok(Object.values(config.syncPolicy.categories).every(Boolean));
  assert.equal(config.syncPolicy.includeDescription, true);
  assert.equal(config.syncPolicy.includeLocation, true);
  assert.equal(config.syncPolicy.includeSchoolboxLink, true);
  assert.equal(mapping.googleUserId, "existing-user");
  assert.equal(mapping.syncEnabled, true);
  assert.equal(mapping.directoryActive, true);
  assert.equal((await storage.getEventMappings("existing-user"))[0]?.googleEventId, "legacy-event");
});

test("legacy global email uniqueness is replaced by active-directory uniqueness", async () => {
  await storage.discoverUserMappings([{
    googleUserId: "replacement-user",
    googleEmail: "existing@example.edu",
    schoolboxUserId: 84,
    schoolboxEmail: "existing@example.edu",
    displayName: "Replacement User",
    role: "Staff",
    status: "pending",
    lastSyncAt: null,
    lastError: null,
    eventCount: 0,
    updatedAt: new Date().toISOString(),
  }], false);

  const [active] = await storage.listUserMappings();
  assert.equal(active.googleUserId, "replacement-user");
  assert.equal(active.directoryActive, true);

  const allRows = await storage.listUserMappings(undefined, true);
  const legacyRow = allRows.find((mapping) => mapping.googleUserId === "existing-user");
  assert.equal(legacyRow?.directoryActive, false);
  assert.equal(legacyRow?.syncEnabled, true);
  assert.equal(allRows.filter((mapping) => mapping.googleEmail === "existing@example.edu").length, 2);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
