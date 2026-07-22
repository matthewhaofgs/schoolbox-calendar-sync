import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-diagnostics-"));
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
  syncNewUsersByDefault: true,
}, "test:setup");

test("a failed event preserves admin-safe run, user, and event diagnostic context", async () => {
  const run = await runFullSync("test", "test:runner", {
    schoolbox: {
      async getAllUsers() {
        return [{ id: 42, email: "diagnostic.user@example.edu", fullName: "Diagnostic User", enabled: true }];
      },
      async getCalendarEvents() {
        return [{
          sourceKey: "schoolbox:event:42",
          title: "Synthetic diagnostic event",
          description: "A synthetic description used only by the test suite.",
          location: "Test room",
          author: "Test author",
          start: "2026-07-22T09:00:00+10:00",
          end: "2026-07-22T10:00:00+10:00",
          allDay: false,
          type: "Test lesson",
          category: "timetable",
          sourceUrl: "https://schoolbox.example.edu/calendar/event/42",
        }];
      },
    },
    google: {
      async listAllUsers() {
        return [{ id: "google-diagnostic-user", primaryEmail: "diagnostic.user@example.edu", suspended: false }];
      },
      async insertEvent() {
        throw new Error("Synthetic Calendar API failure");
      },
    },
  });

  assert.equal(run.status, "completed_with_errors");
  assert.equal(run.errors, 1);

  const users = await storage.listRunUserDiagnostics(run.id);
  assert.equal(users.length, 1);
  assert.equal(users[0].status, "failed");
  assert.equal(users[0].stage, "processing_events");
  assert.equal(users[0].eventsFound, 1);
  assert.equal(users[0].eventsIncluded, 1);
  assert.match(users[0].errorMessage ?? "", /Synthetic Calendar API failure/);

  const eventResult = await storage.listRunEventDiagnostics(run.id, "google-diagnostic-user");
  assert.equal(eventResult.total, 1);
  assert.equal(eventResult.events[0].action, "failed");
  assert.equal(eventResult.events[0].title, "Synthetic diagnostic event");
  assert.equal(eventResult.events[0].description, "A synthetic description used only by the test suite.");
  assert.equal(eventResult.events[0].location, "Test room");
  assert.equal(eventResult.events[0].author, "Test author");
  assert.equal(eventResult.events[0].eventType, "Test lesson");
  assert.equal(eventResult.events[0].category, "timetable");
  assert.match(eventResult.events[0].errorMessage ?? "", /Synthetic Calendar API failure/);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
