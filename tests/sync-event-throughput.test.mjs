import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

const temporary = mkdtempSync(join(tmpdir(), "relay-sync-event-throughput-"));
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
await completeStoredSetup(storage);

const sourceEvents = Array.from({ length: 18 }, (_, index) => ({
  sourceKey: `lesson-${index}`,
  title: `Lesson ${index}`,
  description: "",
  location: null,
  start: `2026-07-24T${String(8 + Math.floor(index / 2)).padStart(2, "0")}:${index % 2 ? "30" : "00"}:00+10:00`,
  end: `2026-07-24T${String(8 + Math.floor(index / 2)).padStart(2, "0")}:${index % 2 ? "59" : "29"}:00+10:00`,
  allDay: false,
  type: "Timetable",
  category: "timetable",
  completed: false,
}));

let activeDeletes = 0;
let peakDeletes = 0;
const clients = {
  schoolbox: {
    async getAllUsers() {
      return [{ id: 101, email: "pilot@example.edu", enabled: true }];
    },
    async getCalendarEvents() {
      return sourceEvents;
    },
  },
  google: {
    async listAllUsers() {
      return [{ id: "google-pilot", primaryEmail: "pilot@example.edu", suspended: false }];
    },
    async insertEvent() {},
    async updateEvent() {},
    async deleteEvent() {
      activeDeletes += 1;
      peakDeletes = Math.max(peakDeletes, activeDeletes);
      try {
        await new Promise(resolve => setTimeout(resolve, 40));
      } finally {
        activeDeletes -= 1;
      }
    },
  },
};

test("a large exclusion reconciliation completes within the per-user deadline", async () => {
  await runFullSync("test", "test:runner", clients);
  await storage.setUsersSyncEnabled(["google-pilot"], true, "test:administrator");
  const baseline = await runFullSync("test", "test:runner", clients);
  assert.equal(baseline.eventsCreated, sourceEvents.length);

  await storage.saveUserEventExclusions("google-pilot", {
    categories: ["timetable"],
    eventTypes: [],
  }, "test:administrator");

  const restricted = await runFullSync("test", "test:runner", clients, {
    discoveryTimeoutMs: 200,
    userSyncTimeoutMs: 300,
    runTimeoutMs: 1_000,
  });

  assert.equal(restricted.status, "completed");
  assert.equal(restricted.eventsDeleted, sourceEvents.length);
  assert.equal((await storage.getEventMappings("google-pilot")).length, 0);
  assert.equal(peakDeletes, 6, "Google event operations should use the bounded per-user pool");
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
