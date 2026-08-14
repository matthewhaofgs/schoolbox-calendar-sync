import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

const temporary = mkdtempSync(join(tmpdir(), "relay-sync-timeout-"));
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
await completeStoredSetup(storage);

function abortAwareNever(signal) {
  return new Promise((_resolve, reject) => {
    const aborted = () => reject(signal.reason ?? new Error("aborted"));
    if (signal.aborted) aborted();
    else signal.addEventListener("abort", aborted, { once: true });
  });
}

test("initial discovery aborts and fails the run instead of retaining a live heartbeat", async () => {
  const started = Date.now();
  const run = await runFullSync("test", "test:runner", {
    schoolbox: {
      async getAllUsers({ signal }) { return abortAwareNever(signal); },
      async getCalendarEvents() { return []; },
    },
    google: {
      async listAllUsers(_admin, { signal }) { return abortAwareNever(signal); },
    },
  }, {
    discoveryTimeoutMs: 20,
    userSyncTimeoutMs: 100,
    runTimeoutMs: 500,
  });

  assert.equal(run.status, "failed");
  assert.match(run.message ?? "", /Schoolbox user discovery timed out/);
  assert.ok(Date.now() - started < 1_000, "the timed-out run should settle promptly");
  assert.ok(run.completedAt);
  assert.notEqual(run.startedAt, run.completedAt);
  assert.equal((await storage.listRuns(1))[0].status, "failed");
});

test("the whole-run limit settles even when an injected dependency ignores cancellation", async () => {
  const run = await runFullSync("test", "test:runner", {
    schoolbox: {
      async getAllUsers() { return new Promise(() => undefined); },
      async getCalendarEvents() { return []; },
    },
    google: {
      async listAllUsers() { return []; },
    },
  }, {
    discoveryTimeoutMs: 500,
    userSyncTimeoutMs: 100,
    runTimeoutMs: 20,
  });

  assert.equal(run.status, "failed");
  assert.match(run.message ?? "", /Organization synchronization timed out/);
  assert.ok(run.completedAt);
});

test("one stalled user becomes a user error while the organization run completes", async () => {
  const now = new Date().toISOString();
  await storage.upsertEventMapping({
    googleUserId: "google-pilot",
    sourceKey: "existing-managed-event",
    googleEventId: "google-event-1",
    calendarId: "primary",
    sourceHash: "existing-hash",
    sourceStart: "2026-07-23T09:00:00.000Z",
    sourceEnd: "2026-07-23T09:30:00.000Z",
    lastSeenRunId: "earlier-run",
    createdAt: now,
    updatedAt: now,
  });

  const run = await runFullSync("test", "test:runner", {
    schoolbox: {
      async getAllUsers() {
        return [{ id: 101, email: "pilot@example.edu", enabled: true }];
      },
      async getCalendarEvents(_userId, { signal }) {
        return abortAwareNever(signal);
      },
    },
    google: {
      async listAllUsers() {
        return [{ id: "google-pilot", primaryEmail: "pilot@example.edu", suspended: false }];
      },
    },
  }, {
    discoveryTimeoutMs: 100,
    userSyncTimeoutMs: 20,
    runTimeoutMs: 500,
  });

  assert.equal(run.status, "completed_with_errors");
  assert.equal(run.usersDiscovered, 1);
  assert.equal(run.usersMatched, 1);
  assert.equal(run.usersSynced, 0);
  assert.equal(run.errors, 1);
  const mapping = await storage.getUserMapping("google-pilot");
  assert.equal(mapping?.status, "error");
  assert.match(mapping?.lastError ?? "", /User calendar synchronization timed out/);
  assert.equal(mapping?.eventCount, 1, "partial progress should preserve the actual managed event count");
  const outcomes = await storage.listRunUserDiagnostics(run.id);
  assert.equal(outcomes.length, 1);
  assert.equal(outcomes[0].status, "failed");
  assert.equal(outcomes[0].stage, "fetching_events");
  assert.equal(outcomes[0].managedEventsAfter, 1);
  assert.match(outcomes[0].errorMessage ?? "", /User calendar synchronization timed out/);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
