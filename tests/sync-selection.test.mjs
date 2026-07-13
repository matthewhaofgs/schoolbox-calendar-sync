import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-sync-selection-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { cleanupUserManagedEvents, runFullSync } = await import("../lib/sync.ts");
const { db } = await import("../lib/db.ts");

await storage.saveConfig({
  schoolboxBaseUrl: "https://schoolbox.example.edu",
  schoolboxToken: "test-schoolbox-token",
  googleAdminEmail: "workspace-admin@example.edu",
  googleServiceAccountJson: JSON.stringify({
    type: "service_account",
    client_email: "relay@example-project.iam.gserviceaccount.com",
    client_id: "1234567890",
    private_key: "unused-by-injected-test-client",
  }),
  syncNewUsersByDefault: false,
}, "test:setup");

const schoolboxUsers = [
  { id: 101, email: "enabled@example.edu", fullName: "Enabled User", enabled: true },
  { id: 202, email: "paused@example.edu", fullName: "Paused User", enabled: true },
];
const googleUsers = [
  { id: "google-enabled", primaryEmail: "enabled@example.edu", suspended: false },
  { id: "google-paused", primaryEmail: "paused@example.edu", suspended: false },
];
const calls = {
  schoolboxCalendarUsers: [],
  insertedFor: [],
  updatedFor: [],
  deletedFor: [],
};

const clients = {
  schoolbox: {
    async getAllUsers() {
      return schoolboxUsers;
    },
    async getCalendarEvents(userId) {
      calls.schoolboxCalendarUsers.push(userId);
      return [{
        sourceKey: `event-${userId}`,
        title: `Calendar event for ${userId}`,
        description: "Selection safety test",
        location: null,
        start: "2026-07-14T09:00:00+10:00",
        end: "2026-07-14T10:00:00+10:00",
        allDay: false,
        type: "event",
      }];
    },
  },
  google: {
    async listAllUsers() {
      return googleUsers;
    },
    async insertEvent(userEmail) {
      calls.insertedFor.push(userEmail);
    },
    async updateEvent(userEmail) {
      calls.updatedFor.push(userEmail);
    },
    async deleteEvent(userEmail) {
      calls.deletedFor.push(userEmail);
    },
  },
};

test("runFullSync never processes paused matches and does process enabled matches", async () => {
  const discoveryRun = await runFullSync("test", "test:runner", clients);
  assert.equal(discoveryRun.usersMatched, 2);
  assert.equal(discoveryRun.usersSynced, 0);
  assert.deepEqual(calls.schoolboxCalendarUsers, []);
  assert.deepEqual(calls.insertedFor, []);
  assert.deepEqual(calls.updatedFor, []);
  assert.deepEqual(calls.deletedFor, []);

  await storage.setUsersSyncEnabled(["google-enabled"], true, "local:administrator");
  const enabledRun = await runFullSync("test", "test:runner", clients);

  assert.equal(enabledRun.usersMatched, 2);
  assert.equal(enabledRun.usersSynced, 1);
  assert.equal(enabledRun.eventsCreated, 1);
  assert.deepEqual(calls.schoolboxCalendarUsers, [101]);
  assert.deepEqual(calls.insertedFor, ["enabled@example.edu"]);
  assert.deepEqual(calls.updatedFor, []);
  assert.deepEqual(calls.deletedFor, []);

  const mappings = new Map((await storage.listUserMappings()).map((mapping) => [mapping.googleUserId, mapping]));
  assert.equal(mappings.get("google-enabled")?.status, "synced");
  assert.equal(mappings.get("google-paused")?.syncEnabled, false);
  assert.equal(mappings.get("google-paused")?.lastSyncAt, null);

  const cleanup = await cleanupUserManagedEvents(
    "google-enabled",
    "local:administrator",
    clients.google,
  );
  assert.deepEqual(cleanup, {
    paused: true,
    deleted: 1,
    alreadyMissing: 0,
    remaining: 0,
    error: null,
  });
  assert.deepEqual(calls.deletedFor, ["enabled@example.edu"]);
  assert.equal((await storage.getEventMappings("google-enabled")).length, 0);
  const cleanedUser = await storage.getUserMapping("google-enabled");
  assert.equal(cleanedUser?.syncEnabled, false, "cleanup must pause the user before removing events");
  assert.equal(cleanedUser?.eventCount, 0);
  assert.equal(cleanedUser?.status, "pending");

  await storage.setUsersSyncEnabled(["google-enabled"], true, "local:administrator");
  await runFullSync("test", "test:runner", clients);
  const failedCleanup = await cleanupUserManagedEvents(
    "google-enabled",
    "local:administrator",
    { async deleteEvent() { throw new Error("simulated Google delete failure"); } },
  );
  assert.equal(failedCleanup.deleted, 0);
  assert.equal(failedCleanup.remaining, 1, "a failed Google deletion must keep Relay's mapping for retry");
  assert.match(failedCleanup.error ?? "", /simulated Google delete failure/);
  assert.equal((await storage.getEventMappings("google-enabled")).length, 1);
  const cleanupFailureUser = await storage.getUserMapping("google-enabled");
  assert.equal(cleanupFailureUser?.syncEnabled, false);
  assert.equal(cleanupFailureUser?.status, "error");
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
