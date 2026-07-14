import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-user-selection-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { db } = await import("../lib/db.ts");

await storage.ensureSchema();

function discovery(googleUserId, googleEmail, overrides = {}) {
  return {
    googleUserId,
    googleEmail,
    schoolboxUserId: Number(googleUserId.replace(/\D/g, "")) || 1,
    schoolboxEmail: googleEmail,
    displayName: googleEmail.split("@", 1)[0],
    role: "Student",
    status: "matched",
    lastSyncAt: null,
    lastError: null,
    eventCount: 0,
    updatedAt: new Date().toISOString(),
    ...overrides,
  };
}

test("fresh installations leave newly discovered users paused by default", async () => {
  const config = await storage.getConfig();
  assert.equal(config.syncNewUsersByDefault, false);
  assert.equal(config.syncPolicy.eventTypeMode, "all");
  assert.ok(Object.values(config.syncPolicy.categories).every(Boolean));
  assert.equal(config.syncPolicy.deleteMissingEvents, true);
  assert.equal(config.syncPolicy.deleteExcludedEvents, true);

  await storage.discoverUserMappings(
    [discovery("google-1", "pilot.one@example.edu")],
    config.syncNewUsersByDefault,
  );

  const [mapping] = await storage.listUserMappings();
  assert.equal(mapping.googleUserId, "google-1");
  assert.equal(mapping.syncEnabled, false);
});

test("configuration rejects an invalid calendar time zone", async () => {
  await assert.rejects(
    storage.saveConfig({ timezone: "Not/A_Real_Zone" }, "local:administrator"),
    /valid IANA calendar time zone/i,
  );
});

test("the discovery default is applied only on insert", async () => {
  await storage.setUsersSyncEnabled(["google-1"], true, "local:administrator");

  const changedConfig = await storage.saveConfig(
    { syncNewUsersByDefault: true },
    "local:administrator",
  );
  assert.equal(changedConfig.syncNewUsersByDefault, true);

  await storage.discoverUserMappings(
    [discovery("google-1", "renamed.pilot@example.edu", { displayName: "Renamed Pilot" })],
    false,
  );

  const existing = (await storage.listUserMappings()).find((mapping) => mapping.googleUserId === "google-1");
  assert.equal(existing.googleEmail, "renamed.pilot@example.edu");
  assert.equal(existing.displayName, "Renamed Pilot");
  assert.equal(existing.syncEnabled, true, "rediscovery must preserve the administrator's explicit selection");

  await storage.discoverUserMappings(
    [discovery("google-2", "automatic@example.edu")],
    changedConfig.syncNewUsersByDefault,
  );
  const inserted = (await storage.listUserMappings()).find((mapping) => mapping.googleUserId === "google-2");
  assert.equal(inserted.syncEnabled, true);
});

test("administrators can enable and pause multiple users in one operation", async () => {
  await storage.discoverUserMappings(
    [
      discovery("google-3", "pilot.three@example.edu"),
      discovery("google-4", "pilot.four@example.edu"),
    ],
    false,
  );

  await storage.setUsersSyncEnabled(["google-3", "google-4"], true, "local:administrator");
  let selected = new Map((await storage.listUserMappings()).map((mapping) => [mapping.googleUserId, mapping.syncEnabled]));
  assert.equal(selected.get("google-3"), true);
  assert.equal(selected.get("google-4"), true);

  await storage.setUsersSyncEnabled(["google-3", "google-4"], false, "local:administrator");
  selected = new Map((await storage.listUserMappings()).map((mapping) => [mapping.googleUserId, mapping.syncEnabled]));
  assert.equal(selected.get("google-3"), false);
  assert.equal(selected.get("google-4"), false);

  await assert.rejects(
    storage.setUsersSyncEnabled(["google-3", "missing-user"], true, "local:administrator"),
    /no longer available/i,
  );
  selected = new Map((await storage.listUserMappings()).map((mapping) => [mapping.googleUserId, mapping.syncEnabled]));
  assert.equal(selected.get("google-3"), false, "a rejected bulk update must not partially change valid users");

  const audit = db().prepare("SELECT action, actor FROM audit_log WHERE action = 'users.sync_selection_updated' ORDER BY id DESC LIMIT 1")
    .first();
  assert.equal(audit?.actor, "local:administrator");
});

test("email swaps preserve selections and event mappings by stable Google ID", async () => {
  await storage.discoverUserMappings(
    [
      discovery("swap-a", "alpha@example.edu"),
      discovery("swap-b", "beta@example.edu"),
    ],
    false,
  );
  await storage.setUsersSyncEnabled(["swap-a"], true, "local:administrator");

  const timestamp = new Date().toISOString();
  await storage.upsertEventMapping({
    googleUserId: "swap-a",
    sourceKey: "source-a",
    googleEventId: "event-a",
    sourceHash: "hash-a",
    sourceStart: timestamp,
    sourceEnd: new Date(Date.now() + 60_000).toISOString(),
    lastSeenRunId: "run-a",
    createdAt: timestamp,
    updatedAt: timestamp,
  });

  // Deliberately use the order that used to collide with the unique email on
  // swap-a before swap-b had been updated.
  await storage.discoverUserMappings(
    [
      discovery("swap-a", "beta@example.edu"),
      discovery("swap-b", "alpha@example.edu"),
    ],
    true,
  );

  const byId = new Map((await storage.listUserMappings()).map((mapping) => [mapping.googleUserId, mapping]));
  assert.equal(byId.get("swap-a")?.googleEmail, "beta@example.edu");
  assert.equal(byId.get("swap-b")?.googleEmail, "alpha@example.edu");
  assert.equal(byId.get("swap-a")?.syncEnabled, true);
  assert.equal(byId.get("swap-b")?.syncEnabled, false, "rediscovery defaults must not change an existing choice");
  assert.equal((await storage.getEventMappings("swap-a"))[0]?.googleEventId, "event-a");
});

test("a reassigned email activates the new ID and safely retires the stale row", async () => {
  await storage.discoverUserMappings(
    [discovery("departed-id", "reassigned@example.edu")],
    true,
  );
  const timestamp = new Date().toISOString();
  await storage.upsertEventMapping({
    googleUserId: "departed-id",
    sourceKey: "departed-source",
    googleEventId: "departed-event",
    sourceHash: "departed-hash",
    sourceStart: timestamp,
    sourceEnd: new Date(Date.now() + 60_000).toISOString(),
    lastSeenRunId: "departed-run",
    createdAt: timestamp,
    updatedAt: timestamp,
  });

  await storage.discoverUserMappings(
    [discovery("replacement-id", "reassigned@example.edu")],
    false,
  );

  const active = await storage.listUserMappings();
  assert.deepEqual(active.map((mapping) => mapping.googleUserId), ["replacement-id"]);
  assert.equal(active[0]?.syncEnabled, false);
  assert.equal(active[0]?.directoryActive, true);

  const allRows = await storage.listUserMappings(undefined, true);
  const departed = allRows.find((mapping) => mapping.googleUserId === "departed-id");
  assert.equal(departed?.googleEmail, "reassigned@example.edu");
  assert.equal(departed?.syncEnabled, true);
  assert.equal(departed?.directoryActive, false);
  assert.equal((await storage.getEventMappings("departed-id"))[0]?.googleEventId, "departed-event");
  await assert.rejects(
    storage.setUsersSyncEnabled(["departed-id"], false, "local:administrator"),
    /no longer available/i,
  );

  const snapshot = await storage.statusSnapshot();
  assert.equal(snapshot.counts.users, 1);
  assert.equal(snapshot.counts.disabled, 1);
  assert.equal(snapshot.counts.events, 0, "retained stale mappings must not inflate active coverage metrics");
});

test("paused unmatched accounts remain informational in aggregate counts", async () => {
  await storage.discoverUserMappings([
    discovery("google-only", "google-only@example.edu", {
      schoolboxUserId: null,
      schoolboxEmail: null,
      status: "unmatched",
      lastError: "No active Schoolbox user has this primary email address.",
    }),
  ], false);
  const snapshot = await storage.statusSnapshot();
  assert.equal(snapshot.counts.unmatched, 1);
  assert.equal(snapshot.counts.enabled, 0);
  assert.equal(snapshot.counts.errors, 0);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
