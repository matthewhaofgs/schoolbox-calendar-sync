import assert from "node:assert/strict";
import { after, test } from "node:test";
import { randomBytes } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

const temporary = mkdtempSync(join(tmpdir(), "relay-microsoft-duplicate-identity-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const storage = await import("../lib/storage.ts");
const { runFullSync } = await import("../lib/sync.ts");
const { db } = await import("../lib/db.ts");

await storage.saveConfig({
  schoolboxBaseUrl: "https://schoolbox.example.edu",
  schoolboxToken: "synthetic-schoolbox-token",
  googleEnabled: false,
  microsoftEnabled: false,
  microsoftTenantId: "11111111-1111-4111-8111-111111111111",
  microsoftClientId: "22222222-2222-4222-8222-222222222222",
  microsoftClientSecret: "synthetic-client-secret",
  syncNewMicrosoftUsersByDefault: false,
}, "test:setup");
await completeStoredSetup(storage, { google: false, microsoft: true });

const neverWrite = async () => {
  throw new Error("An ambiguous Microsoft identity must never reach calendar synchronization.");
};

const clients = {
  schoolbox: {
    async getAllUsers() {
      return [{ id: 701, email: "shared-address@example.edu", fullName: "Synthetic User", enabled: true }];
    },
    async getCalendarEvents() {
      throw new Error("An ambiguous Microsoft identity must not fetch Schoolbox events.");
    },
  },
  microsoft: {
    async listAllUsers() {
      return [
        {
          id: "entra-object-a",
          displayName: "Synthetic A",
          mail: "shared-address@example.edu",
          userPrincipalName: "synthetic-a@example.onmicrosoft.com",
          proxyAddresses: ["SMTP:shared-address@example.edu"],
          accountEnabled: true,
          userType: "Member",
        },
        {
          id: "entra-object-b",
          displayName: "Synthetic B",
          mail: "shared-address@example.edu",
          userPrincipalName: "synthetic-b@example.onmicrosoft.com",
          proxyAddresses: ["SMTP:shared-address@example.edu"],
          accountEnabled: true,
          userType: "Member",
        },
      ];
    },
    createCalendar: neverWrite,
    updateCalendar: neverWrite,
    insertEvent: neverWrite,
    updateEvent: neverWrite,
    deleteEvent: neverWrite,
  },
};

test("duplicate Entra mail values remain safely unmatched without failing discovery", async () => {
  const run = await runFullSync("test", "test:runner", clients, {}, ["microsoft"]);
  assert.equal(run.status, "completed");
  assert.equal(run.errors, 0);
  assert.equal(run.usersDiscovered, 2);
  assert.equal(run.usersMatched, 0);
  assert.equal(run.usersSynced, 0);

  const [targetRun] = await storage.listRunTargets(run.id);
  assert.equal(targetRun.status, "completed");
  assert.equal(targetRun.usersDiscovered, 2);
  assert.equal(targetRun.usersMatched, 0);
  assert.equal(targetRun.usersSelected, 0);

  const mappings = await storage.listUserMappings(undefined, false, "microsoft");
  assert.equal(mappings.length, 2);
  assert.ok(mappings.every((mapping) => mapping.status === "unmatched"));
  assert.ok(mappings.every((mapping) => mapping.syncEnabled === false));
  assert.ok(mappings.every((mapping) => /shared by multiple active Microsoft 365 accounts/.test(mapping.lastError)));
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
