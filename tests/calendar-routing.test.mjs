import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-calendar-routing-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { retireCalendarDestination, runFullSync } = await import("../lib/sync.ts");
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
  syncPolicy: {
    secondaryCalendars: [{ id: "learning", name: "Learning", description: "Managed learning events" }],
    eventTypeOverrides: {
      lesson: { destinationId: "learning", transparency: "transparent", colorId: "9" },
    },
  },
}, "test:setup");

const calls = { calendarsCreated: [], calendarsUpdated: [], calendarsDeleted: [], inserted: [], updated: [], deleted: [] };
const clients = {
  schoolbox: {
    async getAllUsers() { return [{ id: 101, email: "pilot@example.edu", enabled: true }]; },
    async getCalendarEvents() {
      return [{
        sourceKey: "lesson",
        title: "Lesson",
        description: "",
        location: null,
        start: "2026-07-14T09:00:00+10:00",
        end: "2026-07-14T10:00:00+10:00",
        allDay: false,
        type: "Lesson",
        category: "timetable",
        completed: false,
      }];
    },
  },
  google: {
    async listAllUsers() { return [{ id: "google-pilot", primaryEmail: "pilot@example.edu", suspended: false }]; },
    async createCalendar(_email, calendar) {
      calls.calendarsCreated.push(calendar);
      return { id: "google-secondary-calendar" };
    },
    async updateCalendar(_email, calendarId, calendar) { calls.calendarsUpdated.push({ calendarId, calendar }); },
    async deleteCalendar(_email, calendarId) { calls.calendarsDeleted.push(calendarId); },
    async insertEvent(_email, event, options) { calls.inserted.push({ event, options }); },
    async updateEvent(_email, eventId, event, options) { calls.updated.push({ eventId, event, options }); },
    async deleteEvent(_email, eventId, options) { calls.deleted.push({ eventId, options }); },
  },
};

test("per-type routing lazily creates a secondary calendar and safely moves managed events", async () => {
  await runFullSync("test", "test:runner", clients);
  assert.equal(calls.calendarsCreated.length, 0, "paused discovery must not create user calendars");

  await storage.setUsersSyncEnabled(["google-pilot"], true, "test:administrator");
  const secondaryRun = await runFullSync("test", "test:runner", clients);
  assert.equal(secondaryRun.eventsCreated, 1);
  assert.equal(calls.calendarsCreated.length, 1);
  assert.equal(calls.inserted[0].options.calendarId, "google-secondary-calendar");
  assert.equal(calls.inserted[0].event.transparency, "transparent");
  assert.equal(calls.inserted[0].event.colorId, "9");
  assert.equal((await storage.getEventMappings("google-pilot"))[0].calendarId, "google-secondary-calendar");

  await storage.saveConfig({
    syncPolicy: {
      secondaryCalendars: [{ id: "learning", name: "Learning Hub", description: "Renamed managed learning events" }],
      eventTypeOverrides: { lesson: { destinationId: "primary", transparency: "opaque", colorId: "" } },
    },
  }, "test:administrator");
  const movedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(movedRun.eventsUpdated, 1);
  assert.deepEqual(calls.calendarsUpdated, [{
    calendarId: "google-secondary-calendar",
    calendar: {
      summary: "Learning Hub",
      description: "Renamed managed learning events",
      timeZone: "Australia/Sydney",
    },
  }], "an existing destination is renamed even when this run routes no events to it");
  assert.equal((await storage.getUserCalendarTarget("google-pilot", "learning")).summary, "Learning Hub");
  assert.equal(calls.inserted.at(-1).options.calendarId, "primary", "the new copy is written before the old copy is deleted");
  assert.equal(calls.deleted.at(-1).options.calendarId, "google-secondary-calendar");
  assert.equal((await storage.getEventMappings("google-pilot"))[0].calendarId, "primary");

  const retirement = await retireCalendarDestination("learning", "test:administrator", clients.google);
  assert.deepEqual(retirement, {
    destinationId: "learning",
    calendarsDeleted: 1,
    calendarsAlreadyMissing: 0,
    calendarsFailed: 0,
    calendarsRemaining: 0,
    eventMappingsRemoved: 0,
    error: null,
  });
  assert.deepEqual(calls.calendarsDeleted, ["google-secondary-calendar"]);
  assert.equal((await storage.getConfig(false)).syncPolicy.secondaryCalendars.length, 0);
  assert.equal(await storage.getUserCalendarTarget("google-pilot", "learning"), null);
  assert.equal((await storage.getEventMappings("google-pilot"))[0].calendarId, "primary", "retiring a secondary destination must not remove primary-calendar mappings");

  await storage.saveConfig({
    syncPolicy: {
      secondaryCalendars: [{ id: "archive", name: "Archive", description: "Retirement retry" }],
    },
  }, "test:administrator");
  const timestamp = new Date().toISOString();
  await storage.upsertUserCalendarTarget({
    googleUserId: "google-pilot",
    destinationId: "archive",
    googleCalendarId: "google-archive-calendar",
    summary: "Archive",
    description: "Retirement retry",
    timeZone: "Australia/Sydney",
    createdAt: timestamp,
    updatedAt: timestamp,
  });
  await storage.upsertEventMapping({
    googleUserId: "google-pilot",
    sourceKey: "archive-event",
    googleEventId: "google-archive-event",
    calendarId: "google-archive-calendar",
    sourceHash: "archive-hash",
    sourceStart: timestamp,
    sourceEnd: new Date(Date.parse(timestamp) + 30 * 60_000).toISOString(),
    lastSeenRunId: "archive-run",
    createdAt: timestamp,
    updatedAt: timestamp,
  });
  const partial = await retireCalendarDestination("archive", "test:administrator", {
    async deleteCalendar() { throw new Error("simulated calendar deletion failure"); },
  });
  assert.equal(partial.calendarsFailed, 1);
  assert.equal(partial.calendarsRemaining, 1);
  assert.match(partial.error ?? "", /simulated calendar deletion failure/);
  assert.equal((await storage.getConfig(false)).syncPolicy.secondaryCalendars.length, 0, "routing is removed before deletion so a sync cannot recreate a partly retired destination");
  assert.equal((await storage.listCalendarDestinationUsage())[0].destinationId, "archive", "failed targets remain visible for retry");

  const retry = await retireCalendarDestination("archive", "test:administrator", {
    async deleteCalendar() {},
  });
  assert.equal(retry.calendarsDeleted, 1);
  assert.equal(retry.calendarsRemaining, 0);
  assert.equal(retry.eventMappingsRemoved, 1);
  assert.equal((await storage.getEventMappings("google-pilot")).some(mapping => mapping.sourceKey === "archive-event"), false);
  assert.deepEqual(await storage.listCalendarDestinationUsage(), []);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
