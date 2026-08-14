import assert from "node:assert/strict";
import { randomBytes } from "node:crypto";
import { after, test } from "node:test";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";
import { completeStoredSetup } from "./setup-fixtures.mjs";

const temporary = mkdtempSync(join(tmpdir(), "relay-microsoft-email-refresh-"));
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

const directoryUser = {
  id: "stable-entra-object",
  displayName: "Synthetic User",
  mail: "old-address@example.edu",
  userPrincipalName: "old-address@example.edu",
  proxyAddresses: ["SMTP:old-address@example.edu"],
  accountEnabled: true,
  userType: "Member",
};

const clients = {
  schoolbox: {
    async getAllUsers() {
      return [{ id: 801, email: "old-address@example.edu", fullName: "Synthetic User", enabled: true }];
    },
    async getCalendarEvents() {
      throw new Error("A paused discovery test must not fetch Schoolbox events.");
    },
  },
  microsoft: {
    async listAllUsers() { return [directoryUser]; },
  },
};

test("Microsoft rediscovery follows the primary SMTP address while preserving aliases for matching", async () => {
  const first = await runFullSync("test", "test:first", clients, {}, ["microsoft"]);
  assert.equal(first.status, "completed");
  let [mapping] = await storage.listUserMappings(undefined, false, "microsoft");
  assert.equal(mapping.targetEmail, "old-address@example.edu");
  assert.equal(mapping.schoolboxEmail, "old-address@example.edu");

  directoryUser.proxyAddresses = [
    "smtp:old-address@example.edu",
    "SMTP:new-address@example.edu",
  ];
  const second = await runFullSync("test", "test:second", clients, {}, ["microsoft"]);
  assert.equal(second.status, "completed");
  [mapping] = await storage.listUserMappings(undefined, false, "microsoft");
  assert.equal(mapping.targetUserId, "stable-entra-object");
  assert.equal(mapping.targetEmail, "new-address@example.edu");
  assert.equal(mapping.schoolboxEmail, "old-address@example.edu");
  assert.equal(mapping.syncEnabled, false);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
