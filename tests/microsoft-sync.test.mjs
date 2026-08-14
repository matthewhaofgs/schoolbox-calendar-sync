import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

const temporary = mkdtempSync(join(tmpdir(), "relay-microsoft-sync-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { cleanupUserManagedEvents, microsoftEventBody, retireCalendarDestination, runFullSync } = await import("../lib/sync.ts");
const { MicrosoftGraphError } = await import("../lib/microsoft.ts");
const { DEFAULT_SYNC_POLICY } = await import("../lib/policy.ts");
const { db } = await import("../lib/db.ts");

await storage.saveConfig({
  schoolboxBaseUrl: "https://schoolbox.example.edu",
  schoolboxToken: "synthetic-schoolbox-token",
  googleEnabled: true,
  googleAdminEmail: "workspace-admin@example.edu",
  googleServiceAccountJson: JSON.stringify({
    type: "service_account",
    client_email: "relay@example-project.iam.gserviceaccount.com",
    client_id: "1234567890",
    private_key: "unused-by-injected-test-client",
  }),
  microsoftEnabled: false,
  microsoftTenantId: "11111111-1111-4111-8111-111111111111",
  microsoftClientId: "22222222-2222-4222-8222-222222222222",
  microsoftClientSecret: "synthetic-client-secret",
  syncNewGoogleUsersByDefault: false,
  syncNewMicrosoftUsersByDefault: false,
  microsoftSyncPolicy: {
    secondaryCalendars: [{ id: "learning", name: "Learning", description: "Managed learning events" }],
    categoryOverrides: { timetable: { destinationId: "learning", transparency: "opaque" } },
  },
}, "test:setup");
await completeStoredSetup(storage, { microsoft: true });

const eventStart = new Date(Date.now() + 86_400_000);
eventStart.setUTCHours(1, 0, 0, 0);
const eventEnd = new Date(eventStart.getTime() + 60 * 60_000);
let eventTitle = "Synthetic lesson";
let microsoftUpdateNotFoundOnce = false;
const calls = {
  googleListed: 0,
  googleInserted: [],
  microsoftListed: 0,
  microsoftCalendars: [],
  microsoftInserted: [],
  microsoftUpdated: [],
  microsoftDeleted: [],
  microsoftCalendarsDeleted: [],
};

const clients = {
  schoolbox: {
    async getAllUsers() {
      return [{ id: 501, email: "pilot@example.edu", fullName: "Pilot User", enabled: true }];
    },
    async getCalendarEvents() {
      return [{
        sourceKey: "synthetic:lesson",
        title: eventTitle,
        description: "Synthetic test data",
        location: "Test room",
        start: eventStart.toISOString(),
        end: eventEnd.toISOString(),
        allDay: false,
        type: "Lesson",
        category: "timetable",
      }];
    },
  },
  google: {
    async listAllUsers() {
      calls.googleListed += 1;
      return [{ id: "google-pilot", primaryEmail: "pilot@example.edu", suspended: false }];
    },
    async insertEvent(email, event, options) {
      calls.googleInserted.push({ email, event, options });
    },
    async updateEvent() {},
    async deleteEvent() {},
    async createCalendar() { throw new Error("Google should remain on its primary calendar"); },
    async updateCalendar() {},
  },
  microsoft: {
    async listAllUsers() {
      calls.microsoftListed += 1;
      return [{ id: "microsoft-pilot", mail: "pilot@example.edu", displayName: "Pilot User", accountEnabled: true, userType: "Member" }];
    },
    async createCalendar(userId, calendar) {
      calls.microsoftCalendars.push({ userId, calendar });
      return { id: "outlook-learning-calendar" };
    },
    async updateCalendar() {},
    async insertEvent(userId, event, options) {
      calls.microsoftInserted.push({ userId, event, options });
      return { ...event, id: `outlook-event-${calls.microsoftInserted.length}` };
    },
    async updateEvent(userId, eventId, event, options) {
      calls.microsoftUpdated.push({ userId, eventId, event, options });
      if (microsoftUpdateNotFoundOnce) {
        microsoftUpdateNotFoundOnce = false;
        throw new MicrosoftGraphError("The event no longer exists", { status: 404, code: "ErrorItemNotFound" });
      }
      return { ...event, id: eventId };
    },
    async deleteEvent(userId, eventId, options) {
      calls.microsoftDeleted.push({ userId, eventId, options });
    },
    async deleteCalendar(userId, calendarId) {
      calls.microsoftCalendarsDeleted.push({ userId, calendarId });
    },
  },
};

test("Google Workspace and Microsoft 365 discovery, selection, routing, mappings, and cleanup stay independent", async () => {
  await runFullSync("test", "test:runner", clients, {}, ["google", "microsoft"]);
  assert.equal(calls.googleListed, 1);
  assert.equal(calls.microsoftListed, 1);
  assert.equal(calls.googleInserted.length, 0);
  assert.equal(calls.microsoftInserted.length, 0);

  await storage.setUsersSyncEnabled(["microsoft-pilot"], true, "test:administrator", "microsoft");
  const microsoftRun = await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  assert.equal(microsoftRun.usersSynced, 1);
  assert.equal(microsoftRun.eventsCreated, 1);
  assert.equal(calls.googleListed, 1, "a Microsoft-only run must not call Google");
  assert.equal(calls.microsoftCalendars.length, 1);
  assert.equal(calls.microsoftInserted[0].options.calendarId, "outlook-learning-calendar");
  assert.equal(calls.microsoftInserted[0].event.showAs, "busy");
  assert.match(calls.microsoftInserted[0].event.start.dateTime, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/);
  assert.match(calls.microsoftInserted[0].event.transactionId, /^[0-9a-f-]{36}$/i);

  const microsoftMappings = await storage.getEventMappings("microsoft-pilot", "microsoft");
  assert.equal(microsoftMappings.length, 1);
  assert.equal(microsoftMappings[0].targetEventId, "outlook-event-1");
  assert.equal(microsoftMappings[0].calendarId, "outlook-learning-calendar");
  assert.equal((await storage.getEventMappings("google-pilot", "google")).length, 0);
  assert.deepEqual((await storage.listRunTargets(microsoftRun.id)).map((item) => item.target), ["microsoft"]);

  await storage.saveConfig({
    microsoftSyncPolicy: {
      secondaryCalendars: [{ id: "learning", name: "Learning", description: "Managed learning events" }],
      categoryOverrides: { timetable: { destinationId: "learning", transparency: "opaque" } },
      includeLocation: false,
      reminderMode: "custom",
      reminderMinutes: 5,
    },
  }, "test:administrator");
  const clearedValuesRun = await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  assert.equal(clearedValuesRun.eventsUpdated, 1);
  assert.deepEqual(calls.microsoftUpdated.at(-1).event.locations, [],
    "an Outlook PATCH must explicitly clear a location that Relay no longer includes");
  assert.equal("location" in calls.microsoftUpdated.at(-1).event, false);
  assert.equal("transactionId" in calls.microsoftUpdated.at(-1).event, false,
    "the create-only idempotency key must not be sent in an Outlook PATCH");
  assert.equal(calls.microsoftUpdated.at(-1).event.isReminderOn, true);
  assert.equal(calls.microsoftUpdated.at(-1).event.reminderMinutesBeforeStart, 5);

  const originalTransactionId = calls.microsoftInserted[0].event.transactionId;
  await storage.saveConfig({
    microsoftSyncPolicy: {
      secondaryCalendars: [{ id: "learning", name: "Learning", description: "Managed learning events" }],
      categoryOverrides: { timetable: { destinationId: "learning", transparency: "opaque" } },
      includeLocation: false,
      reminderMode: "calendar_default",
    },
  }, "test:administrator");
  const defaultsRun = await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  assert.equal(defaultsRun.eventsUpdated, 1);
  assert.equal(calls.microsoftInserted.length, 2,
    "Outlook defaults require a replacement because omitted PATCH fields preserve stale reminders");
  assert.equal(calls.microsoftInserted[1].event.isReminderOn, undefined);
  assert.equal(calls.microsoftInserted[1].event.reminderMinutesBeforeStart, undefined);
  assert.notEqual(calls.microsoftInserted[1].event.transactionId, originalTransactionId,
    "a replacement POST must have its own stable idempotency key");
  assert.equal((await storage.getEventMappings("microsoft-pilot", "microsoft"))[0].targetEventId, "outlook-event-2");

  await storage.saveConfig({
    microsoftSyncPolicy: {
      secondaryCalendars: [{ id: "learning", name: "Learning", description: "Managed learning events" }],
      categoryOverrides: { timetable: { destinationId: "learning", transparency: "opaque" } },
      includeLocation: false,
      reminderMode: "custom",
      reminderMinutes: 10,
    },
  }, "test:administrator");
  await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  eventTitle = "Synthetic lesson (changed)";
  microsoftUpdateNotFoundOnce = true;
  const missingMappedEventRun = await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  assert.equal(missingMappedEventRun.status, "completed");
  assert.equal(missingMappedEventRun.eventsUpdated, 1);
  assert.equal(calls.microsoftInserted.length, 3,
    "a missing mapped Outlook event is recreated instead of failing every later run");
  assert.equal((await storage.getEventMappings("microsoft-pilot", "microsoft"))[0].targetEventId, "outlook-event-3");

  await storage.setUsersSyncEnabled(["google-pilot"], true, "test:administrator", "google");
  const microsoftListingsBeforeGoogleRun = calls.microsoftListed;
  const googleRun = await runFullSync("test", "test:runner", clients, {}, ["google"]);
  assert.equal(googleRun.usersSynced, 1);
  assert.equal(calls.microsoftListed, microsoftListingsBeforeGoogleRun, "a Google-only run must not call Microsoft");
  assert.equal(calls.googleInserted.length, 1);
  assert.equal(calls.googleInserted[0].options.calendarId, "primary");
  assert.equal((await storage.getEventMappings("google-pilot", "google")).length, 1);

  await storage.saveUserEventExclusions("microsoft-pilot", { categories: ["timetable"] }, "test:administrator", "microsoft");
  assert.deepEqual((await storage.getUserEventExclusions("google-pilot", "google")).categories, ["timetable"],
    "per-person Schoolbox exclusions should apply consistently to both delivery targets");

  const cleanup = await cleanupUserManagedEvents("microsoft-pilot", "test:administrator", clients.microsoft, {
    target: "microsoft",
    deleteCalendars: true,
  });
  assert.equal(cleanup.deleted, 1);
  assert.equal(cleanup.calendarsDeleted, 1);
  assert.deepEqual(calls.microsoftDeleted.map(({ userId, eventId, options }) => ({
    userId,
    eventId,
    calendarId: options.calendarId,
  })), [
    { userId: "microsoft-pilot", eventId: "outlook-event-1", calendarId: "outlook-learning-calendar" },
    { userId: "microsoft-pilot", eventId: "outlook-event-3", calendarId: "outlook-learning-calendar" },
  ]);
  assert.deepEqual(calls.microsoftCalendarsDeleted, [{
    userId: "microsoft-pilot",
    calendarId: "outlook-learning-calendar",
  }]);
  assert.equal((await storage.getEventMappings("google-pilot", "google")).length, 1,
    "Microsoft cleanup must not remove Google-managed events");

  const now = new Date().toISOString();
  await storage.upsertEventMapping({
    target: "microsoft",
    targetUserId: "microsoft-pilot",
    targetEventId: "outlook-hung-event",
    sourceKey: "synthetic:hung:occurrence",
    calendarId: "primary",
    sourceHash: "hung-hash",
    sourceStart: now,
    sourceEnd: now,
    lastSeenRunId: "cleanup-timeout-test",
    createdAt: now,
    updatedAt: now,
  });
  const timedCleanup = await cleanupUserManagedEvents("microsoft-pilot", "test:administrator", {
    async deleteEvent() { return new Promise(() => undefined); },
  }, { target: "microsoft", timeoutMs: 20 });
  assert.match(timedCleanup.error, /cleanup exceeded/i);
  assert.equal(timedCleanup.remaining, 1);
  assert.equal((await storage.getEventMappings("microsoft-pilot", "microsoft"))[0].targetEventId, "outlook-hung-event",
    "a timed-out deletion must retain its mapping for a safe retry");

  await storage.saveConfig({
    microsoftSyncPolicy: {
      secondaryCalendars: [{ id: "archive", name: "Archive", description: "Retirement timeout test" }],
    },
  }, "test:administrator");
  await storage.upsertUserCalendarTarget({
    target: "microsoft",
    targetUserId: "microsoft-pilot",
    targetCalendarId: "outlook-archive-calendar",
    googleUserId: "microsoft-pilot",
    googleCalendarId: "outlook-archive-calendar",
    destinationId: "archive",
    summary: "Archive",
    description: "Retirement timeout test",
    timeZone: "Australia/Sydney",
    createdAt: now,
    updatedAt: now,
  });
  await assert.rejects(
    retireCalendarDestination("archive", "test:administrator", {
      async deleteCalendar() { return new Promise(() => undefined); },
    }, "microsoft", { runTimeoutMs: 20 }),
    /retirement exceeded/i,
  );
  assert.equal((await storage.listCalendarTargetsForDestination("archive", "microsoft")).length, 1,
    "a timed-out calendar deletion must retain its target record for retry");
  const retirementRun = (await storage.listRuns(1))[0];
  assert.equal(retirementRun.status, "failed");
  assert.ok(retirementRun.completedAt, "the retirement run must be finalized even when a Graph request hangs");

  await storage.saveConfig({ microsoftClientSecret: "rotated-synthetic-client-secret" }, "test:administrator");
  const rotated = await storage.getConfig(false);
  assert.equal(rotated.microsoftEnabled, false, "credential rotation must pause Microsoft delivery");
  assert.equal(rotated.microsoftConsentGrantedAt, "", "replacement credentials must be reverified");
});

test("Microsoft event translation uses Outlook wall-clock date-time and target-specific semantics", async () => {
  const body = await microsoftEventBody({
    sourceKey: "translation",
    title: "Translation test",
    description: "Description",
    location: "Room",
    start: "2026-08-15T09:30:00+10:00",
    end: "2026-08-15T10:00:00+10:00",
    allDay: false,
    type: "Event",
    category: "other",
  }, "microsoft-user", "Australia/Sydney", "translation:occurrence", {
    ...DEFAULT_SYNC_POLICY,
    transparency: "transparent",
    visibility: "private",
    reminderMode: "custom",
    reminderMinutes: 15,
  });
  assert.equal(body.start.dateTime, "2026-08-15T09:30:00");
  assert.equal(body.end.dateTime, "2026-08-15T10:00:00");
  assert.equal(body.showAs, "free");
  assert.equal(body.sensitivity, "private");
  assert.equal(body.isReminderOn, true);
  assert.equal(body.reminderMinutesBeforeStart, 15);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
