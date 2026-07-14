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
const { runFullSync } = await import("../lib/sync.ts");
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

const calls = { calendarsCreated: [], calendarsUpdated: [], inserted: [], updated: [], deleted: [] };
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
    syncPolicy: { eventTypeOverrides: { lesson: { destinationId: "primary", transparency: "opaque", colorId: "" } } },
  }, "test:administrator");
  const movedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(movedRun.eventsUpdated, 1);
  assert.equal(calls.inserted.at(-1).options.calendarId, "primary", "the new copy is written before the old copy is deleted");
  assert.equal(calls.deleted.at(-1).options.calendarId, "google-secondary-calendar");
  assert.equal((await storage.getEventMappings("google-pilot"))[0].calendarId, "primary");
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
