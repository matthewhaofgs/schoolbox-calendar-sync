import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-user-exclusions-"));
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
  syncPolicy: { eventTypeMode: "all", deleteExcludedEvents: true },
}, "test:setup");

const schoolboxUsers = [
  { id: 101, email: "custom@example.edu", fullName: "Custom Preferences", enabled: true },
  { id: 202, email: "defaults@example.edu", fullName: "Organisation Defaults", enabled: true },
];
const googleUsers = [
  { id: "google-custom", primaryEmail: "custom@example.edu", suspended: false },
  { id: "google-defaults", primaryEmail: "defaults@example.edu", suspended: false },
];
const sourceEvents = [
  { sourceKey: "lesson", title: "Lesson", description: "", location: null, start: "2026-07-20T09:00:00+10:00", end: "2026-07-20T10:00:00+10:00", allDay: false, type: "Timetable", category: "timetable", completed: false },
  { sourceKey: "junior", title: "Junior event", description: "", location: null, start: "2026-07-21T09:00:00+10:00", end: "2026-07-21T10:00:00+10:00", allDay: false, type: "Junior School Event", category: "school_event", completed: false },
  { sourceKey: "whole", title: "Whole school event", description: "", location: null, start: "2026-07-22T09:00:00+10:00", end: "2026-07-22T10:00:00+10:00", allDay: false, type: "Whole School Event", category: "school_event", completed: false },
];
const calls = { inserted: [], deleted: [] };
let activeDeletes = 0;
let peakDeletes = 0;
const clients = {
  schoolbox: {
    async getAllUsers() { return schoolboxUsers; },
    async getCalendarEvents() { return sourceEvents; },
  },
  google: {
    async listAllUsers() { return googleUsers; },
    async insertEvent(email, event) { calls.inserted.push({ email, title: event.summary }); },
    async updateEvent() {},
    async deleteEvent(email, eventId) {
      activeDeletes += 1;
      peakDeletes = Math.max(peakDeletes, activeDeletes);
      try {
        await new Promise(resolve => setTimeout(resolve, 10));
        calls.deleted.push({ email, eventId });
      } finally {
        activeDeletes -= 1;
      }
    },
  },
};

test("per-user category and exact-type exclusions reconcile only that person's managed events", async () => {
  await runFullSync("test", "test:runner", clients);
  await storage.setUsersSyncEnabled(["google-custom", "google-defaults"], true, "test:administrator");
  const baseline = await runFullSync("test", "test:runner", clients);
  assert.equal(baseline.eventsCreated, 6);
  assert.equal((await storage.getEventMappings("google-custom")).length, 3);
  assert.equal((await storage.getEventMappings("google-defaults")).length, 3);

  const saved = await storage.saveUserEventExclusions("google-custom", {
    categories: ["timetable"],
    eventTypes: ["Junior School Event"],
  }, "test:administrator");
  assert.deepEqual(saved.categories, ["timetable"]);
  assert.deepEqual(saved.eventTypes, ["Junior School Event"]);
  assert.equal((await storage.listUserMappings()).find(user => user.googleUserId === "google-custom")?.hasCustomExclusions, true);

  const restricted = await runFullSync("test", "test:runner", clients);
  assert.equal(restricted.eventsDeleted, 2);
  assert.equal((await storage.getEventMappings("google-custom")).length, 1);
  assert.equal((await storage.getEventMappings("google-defaults")).length, 3);
  assert.ok(calls.deleted.every(call => call.email === "custom@example.edu"), "another person's managed events must not be touched");
  assert.equal(peakDeletes, 2, "independent managed-event removals should run concurrently");

  const diagnostics = await storage.listRunEventDiagnostics(restricted.id, "google-custom");
  const excluded = diagnostics.events.filter(event => event.action === "deleted");
  assert.equal(excluded.length, 2);
  assert.ok(excluded.every(event => /excluded for this person/i.test(event.detail ?? "")));

  const cleared = await storage.saveUserEventExclusions("google-custom", {
    categories: [],
    eventTypes: [],
  }, "test:administrator");
  assert.deepEqual(cleared.categories, []);
  assert.deepEqual(cleared.eventTypes, []);
  assert.equal((await storage.getUserEventExclusions("google-custom")).categories.length, 0);
  assert.equal((await storage.getUserMapping("google-custom"))?.hasCustomExclusions, false);
});

test("unmatched users cannot receive event exclusions", async () => {
  await assert.rejects(
    storage.saveUserEventExclusions("missing-user", { categories: ["timetable"], eventTypes: [] }, "test:administrator"),
    /user not found/i,
  );
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
