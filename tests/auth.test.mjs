import assert from "node:assert/strict";
import { after, test } from "node:test";
import { createHash, pbkdf2Sync, randomBytes, randomUUID } from "node:crypto";
import { mkdtempSync, rmSync } from "node:fs";
import { join } from "node:path";
import { tmpdir } from "node:os";

const temporary = mkdtempSync(join(tmpdir(), "relay-auth-"));
process.env.DATABASE_PATH = join(temporary, "relay.sqlite");
process.env.APP_ORIGIN = "http://127.0.0.1:3000";
process.env.SESSION_SECRET = "test-session-secret-that-is-longer-than-thirty-two-characters";
process.env.SCHEDULER_TOKEN = "test-scheduler-token-that-is-longer-than-thirty-two-characters";
process.env.CONFIG_ENCRYPTION_KEY = randomBytes(32).toString("base64");
process.env.NODE_ENV = "test";

const auth = await import("../lib/auth.ts");
const { db } = await import("../lib/db.ts");

auth.ensureAuthSchema();
const password = "a-strong-test-password";
const salt = randomBytes(16);
const iterations = 600_000;
const passwordHash = `pbkdf2-sha256$${iterations}$${salt.toString("base64url")}$${pbkdf2Sync(password, salt, iterations, 32, "sha256").toString("base64url")}`;
const ownerId = randomUUID();
const now = new Date().toISOString();
db().prepare(`INSERT INTO auth_users
  (id, provider, username, display_name, role, is_owner, enabled, password_hash, created_by, created_at, updated_at)
  VALUES (?, 'local', 'administrator', 'Administrator', 'admin', 1, 1, ?, 'test', ?, ?)`)
  .bind(ownerId, passwordHash, now, now).run();

function cookieFromHeader(header) {
  return header.split(";", 1)[0];
}

function request(path, { method = "GET", cookie = "", csrf = "", origin = process.env.APP_ORIGIN } = {}) {
  const headers = new Headers({ host: "127.0.0.1:3000" });
  if (cookie) headers.set("cookie", cookie);
  if (method !== "GET") {
    headers.set("origin", origin);
    headers.set("sec-fetch-site", origin === process.env.APP_ORIGIN ? "same-origin" : "cross-site");
    if (csrf) headers.set("x-csrf-token", csrf);
  }
  return new Request(`${process.env.APP_ORIGIN}${path}`, { method, headers });
}

test("local administrator login creates an opaque owner session", () => {
  const result = auth.localLogin("administrator", password);
  const cookie = cookieFromHeader(result.cookie);
  assert.equal(result.session.displayName, "Administrator");
  assert.equal(result.session.isOwner, true);
  assert.deepEqual(result.session.permissions, ["view", "operate", "configure", "manage_access"]);
  assert.equal(auth.currentSession(request("/api/status", { cookie }))?.actor, "local:administrator");
});

test("mutating requests require the exact origin and CSRF token", () => {
  const result = auth.localLogin("administrator", password);
  const cookie = cookieFromHeader(result.cookie);
  assert.throws(() => auth.requireSession(request("/api/config", { method: "PUT", cookie }), "configure"), /security token/i);
  assert.throws(() => auth.requireSession(request("/api/config", { method: "PUT", cookie, csrf: result.session.csrfToken, origin: "http://evil.invalid" }), "configure"), /Cross-site|origin/i);
  assert.equal(
    auth.requireSession(request("/api/config", { method: "PUT", cookie, csrf: result.session.csrfToken }), "configure").actor,
    "local:administrator",
  );
});

test("Google staff roles are server enforced and changes revoke sessions", () => {
  const account = auth.saveStaffAccount({ email: "operator@school.edu.au", displayName: "IT Operator", role: "operator", enabled: true }, "local:administrator");
  const rawToken = randomBytes(32).toString("base64url");
  db().prepare("INSERT INTO auth_sessions (token_hash, user_id, created_at, last_seen_at, absolute_expires_at) VALUES (?, ?, ?, ?, ?)")
    .bind(createHash("sha256").update(rawToken).digest("hex"), account.id, now, now, new Date(Date.now() + 3_600_000).toISOString()).run();
  const cookie = `relay_session=${rawToken}`;
  const session = auth.currentSession(request("/api/status", { cookie }));
  assert.equal(session?.role, "operator");
  assert.throws(() => auth.requireSession(request("/api/config", { cookie }), "configure"), /permission/i);
  assert.equal(auth.requireSession(request("/api/status", { cookie }), "view").email, "operator@school.edu.au");

  auth.saveStaffAccount({ ...account, role: "viewer" }, "local:administrator");
  assert.equal(auth.currentSession(request("/api/status", { cookie })), null);

  db().prepare("UPDATE auth_users SET google_sub = 'google-subject-1' WHERE id = ?").bind(account.id).run();
  auth.saveStaffAccount({ ...account, email: "renamed-operator@school.edu.au", role: "viewer" }, "local:administrator");
  const relinked = db().prepare("SELECT email, google_sub FROM auth_users WHERE id = ?").bind(account.id).first();
  assert.equal(relinked.email, "renamed-operator@school.edu.au");
  assert.equal(relinked.google_sub, null);
});

test("Google administrators can manage other administrator accounts", () => {
  const account = auth.saveStaffAccount({ email: "admin@school.edu.au", displayName: "IT Administrator", role: "admin", enabled: true }, "local:administrator");
  const rawToken = randomBytes(32).toString("base64url");
  db().prepare("INSERT INTO auth_sessions (token_hash, user_id, created_at, last_seen_at, absolute_expires_at) VALUES (?, ?, ?, ?, ?)")
    .bind(createHash("sha256").update(rawToken).digest("hex"), account.id, now, now, new Date(Date.now() + 3_600_000).toISOString()).run();
  const session = auth.requireSession(request("/api/admin/staff", { cookie: `relay_session=${rawToken}` }), "manage_access");
  assert.equal(session.role, "admin");
  assert.ok(session.permissions.includes("manage_access"));

  const added = auth.saveStaffAccount({ email: "second-admin@school.edu.au", role: "admin", enabled: true }, session.actor);
  assert.equal(added.role, "admin");
});

test("sessions expose their deadlines and can be explicitly extended", () => {
  const result = auth.localLogin("administrator", password);
  const cookie = cookieFromHeader(result.cookie);
  const rawToken = cookie.split("=", 2)[1];
  const tokenHash = createHash("sha256").update(rawToken).digest("hex");
  const oldLastSeen = new Date(Date.now() - 6 * 60 * 1000).toISOString();
  db().prepare("UPDATE auth_sessions SET last_seen_at = ?, absolute_expires_at = ? WHERE token_hash = ?")
    .bind(oldLastSeen, new Date(Date.now() + 4 * 60 * 1000).toISOString(), tokenHash).run();

  const peeked = auth.currentSession(request("/api/auth/session", { cookie }), { touch: false });
  assert.ok(peeked?.expiresAt);
  assert.ok(peeked?.idleExpiresAt);
  assert.equal(db().prepare("SELECT last_seen_at AS lastSeenAt FROM auth_sessions WHERE token_hash = ?").bind(tokenHash).first()?.lastSeenAt, oldLastSeen);

  const extended = auth.extendSession(request("/api/auth/session", {
    method: "POST",
    cookie,
    csrf: peeked.csrfToken,
  }));
  assert.ok(Date.parse(extended.session.expiresAt) > Date.now() + 7 * 60 * 60 * 1000);
  assert.ok(Date.parse(extended.session.idleExpiresAt) > Date.now() + 29 * 60 * 1000);
});

test("scheduler credentials are isolated from administrator sessions", () => {
  const valid = new Request("http://127.0.0.1:3000/api/sync/local-tick", {
    method: "POST",
    headers: { authorization: `Bearer ${process.env.SCHEDULER_TOKEN}` },
  });
  assert.equal(auth.requestScheduler(valid), "scheduler");
  assert.throws(() => auth.requestScheduler(new Request(valid.url, { method: "POST" })), /authentication failed/i);
});

test("changing the local password revokes sessions and accepts the replacement", () => {
  const result = auth.localLogin("administrator", password);
  const cookie = cookieFromHeader(result.cookie);
  auth.changeLocalPassword(result.session, password, "a-new-strong-test-password");
  assert.equal(auth.currentSession(request("/api/status", { cookie })), null);
  assert.equal(auth.localLogin("administrator", "a-new-strong-test-password").session.isOwner, true);
});

after(() => {
  db().close();
  rmSync(temporary, { recursive: true, force: true });
});
