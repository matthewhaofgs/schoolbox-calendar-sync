import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-sync-policy-"));
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
  syncPolicy: { eventTypeMode: "include", eventTypes: ["Timetable"] },
}, "test:setup");

const calls = { inserted: [], updated: [], deleted: [] };
const sourceEvents = [
  { sourceKey: "lesson", title: "Lesson", description: "", location: null, start: "2026-07-14T09:00:00+10:00", end: "2026-07-14T10:00:00+10:00", allDay: false, type: "Timetable", category: "timetable", completed: false },
  { sourceKey: "excursion", title: "Excursion", description: "", location: null, start: "2026-07-15T09:00:00+10:00", end: "2026-07-15T15:00:00+10:00", allDay: false, type: "Excursion", category: "school_event", completed: false },
];
const clients = {
  schoolbox: {
    async getAllUsers() { return [{ id: 101, email: "pilot@example.edu", enabled: true }]; },
    async getCalendarEvents() { return sourceEvents; },
  },
  google: {
    async listAllUsers() { return [{ id: "google-pilot", primaryEmail: "pilot@example.edu", suspended: false }]; },
    async insertEvent(_email, event) { calls.inserted.push(event.summary); },
    async updateEvent(_email, _id, event) { calls.updated.push(event.summary); },
    async deleteEvent(_email, eventId) { calls.deleted.push(eventId); },
  },
};

test("sync creates only allowed types and reconciles policy exclusions safely", async () => {
  await runFullSync("test", "test:runner", clients);
  await storage.setUsersSyncEnabled(["google-pilot"], true, "test:administrator");

  const timetableRun = await runFullSync("test", "test:runner", clients);
  assert.equal(timetableRun.eventsCreated, 1);
  assert.deepEqual(calls.inserted, ["Lesson"]);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 1);

  await storage.saveConfig({ syncPolicy: { eventTypeMode: "include", eventTypes: ["Excursion"], deleteExcludedEvents: true } }, "test:administrator");
  const switchedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(switchedRun.eventsCreated, 1);
  assert.equal(switchedRun.eventsDeleted, 1);
  assert.deepEqual(calls.inserted, ["Lesson", "Excursion"]);
  assert.equal(calls.deleted.length, 1);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 1);

  await storage.saveConfig({ syncPolicy: { eventTypeMode: "include", eventTypes: ["Timetable"], deleteExcludedEvents: false } }, "test:administrator");
  const retainedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(retainedRun.eventsCreated, 1);
  assert.equal(retainedRun.eventsDeleted, 0);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 2, "excluded managed event should be retained when policy removal is off");
  assert.equal((await storage.getUserMapping("google-pilot"))?.eventCount, 2, "People cleanup controls must count retained managed events");

  await storage.saveConfig({ syncPolicy: { eventTypeMode: "all", deleteMissingEvents: false } }, "test:administrator");
  sourceEvents.splice(1, 1);
  const missingRetainedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(missingRetainedRun.eventsDeleted, 0);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 2);

  await storage.saveConfig({ syncPolicy: { deleteMissingEvents: true } }, "test:administrator");
  const missingRemovedRun = await runFullSync("test", "test:runner", clients);
  assert.equal(missingRemovedRun.eventsDeleted, 1);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 1);

  const catalog = await storage.listDiscoveredEventTypes();
  assert.deepEqual(catalog.map(entry => entry.label), ["Excursion", "Timetable"]);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
