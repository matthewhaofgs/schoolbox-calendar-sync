import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

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
await completeStoredSetup(storage);

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
  deletedCalendarsFor: [],
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
    async deleteCalendar(userEmail, calendarId) {
      calls.deletedCalendarsFor.push({ userEmail, calendarId });
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

  const userOutcomes = await storage.listRunUserDiagnostics(enabledRun.id);
  assert.equal(userOutcomes.length, 1, "paused users should not create run outcome rows");
  assert.equal(userOutcomes[0].status, "completed");
  assert.equal(userOutcomes[0].eventsFound, 1);
  assert.equal(userOutcomes[0].eventsCreated, 1);
  assert.equal(userOutcomes[0].managedEventsAfter, 1);

  const eventOutcomes = await storage.listRunEventDiagnostics(enabledRun.id, "google-enabled");
  assert.equal(eventOutcomes.total, 1);
  assert.equal(eventOutcomes.events[0].action, "created");
  assert.equal(eventOutcomes.events[0].title, "Calendar event for 101");
  assert.equal(eventOutcomes.events[0].eventType, "event");
  assert.equal(eventOutcomes.events[0].category, "other");

  const currentEvents = await storage.getEventMappings("google-enabled");
  assert.equal(currentEvents[0].title, "Calendar event for 101");
  assert.equal(currentEvents[0].destinationId, "primary");

  const targetTimestamp = new Date().toISOString();
  await storage.upsertUserCalendarTarget({
    googleUserId: "google-enabled",
    destinationId: "pilot-events",
    googleCalendarId: "google-pilot-calendar",
    summary: "Pilot events",
    description: "Test calendar",
    timeZone: "Australia/Sydney",
    createdAt: targetTimestamp,
    updatedAt: targetTimestamp,
  });

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
    calendarsDeleted: 0,
    calendarsAlreadyMissing: 0,
    calendarsRemaining: 1,
    error: null,
  });
  assert.deepEqual(calls.deletedFor, ["enabled@example.edu"]);
  assert.equal((await storage.getEventMappings("google-enabled")).length, 0);
  const cleanedUser = await storage.getUserMapping("google-enabled");
  assert.equal(cleanedUser?.syncEnabled, false, "cleanup must pause the user before removing events");
  assert.equal(cleanedUser?.eventCount, 0);
  assert.equal(cleanedUser?.status, "pending");
  assert.equal(cleanedUser?.calendarCount, 1, "event-only cleanup must retain Relay-created calendars");

  const calendarCleanup = await cleanupUserManagedEvents(
    "google-enabled",
    "local:administrator",
    clients.google,
    { deleteCalendars: true },
  );
  assert.equal(calendarCleanup.calendarsDeleted, 1);
  assert.equal(calendarCleanup.calendarsRemaining, 0);
  assert.deepEqual(calls.deletedCalendarsFor, [{
    userEmail: "enabled@example.edu",
    calendarId: "google-pilot-calendar",
  }]);
  assert.equal((await storage.listUserCalendarTargets("google-enabled")).length, 0);

  await storage.setUsersSyncEnabled(["google-enabled"], true, "local:administrator");
  await runFullSync("test", "test:runner", clients);
  await storage.upsertUserCalendarTarget({
    googleUserId: "google-enabled",
    destinationId: "pilot-events",
    googleCalendarId: "google-pilot-calendar-retry",
    summary: "Pilot events",
    description: "Test calendar",
    timeZone: "Australia/Sydney",
    createdAt: targetTimestamp,
    updatedAt: targetTimestamp,
  });
  let calendarDeleteAttempted = false;
  const failedCleanup = await cleanupUserManagedEvents(
    "google-enabled",
    "local:administrator",
    {
      async deleteEvent() { throw new Error("simulated Google delete failure"); },
      async deleteCalendar() { calendarDeleteAttempted = true; },
    },
    { deleteCalendars: true },
  );
  assert.equal(failedCleanup.deleted, 0);
  assert.equal(failedCleanup.remaining, 1, "a failed Google deletion must keep Relay's mapping for retry");
  assert.match(failedCleanup.error ?? "", /simulated Google delete failure/);
  assert.equal(calendarDeleteAttempted, false, "calendar deletion must wait until tracked event cleanup succeeds");
  assert.equal(failedCleanup.calendarsRemaining, 1);
  assert.equal((await storage.getEventMappings("google-enabled")).length, 1);
  const cleanupFailureUser = await storage.getUserMapping("google-enabled");
  assert.equal(cleanupFailureUser?.syncEnabled, false);
  assert.equal(cleanupFailureUser?.status, "error");
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
