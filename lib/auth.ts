import {
  createHash,
  createHmac,
  pbkdf2Sync,
  randomBytes,
  timingSafeEqual,
} from "node:crypto";
import { createRemoteJWKSet, jwtVerify } from "jose";
import { db } from "./db";
import { decryptSecret, encryptSecret, HttpError } from "./security";

export type StaffRole = "admin" | "operator" | "viewer";
export type Permission = "view" | "operate" | "configure" | "manage_access";
export type AuthType = "local" | "google";

export type AuthSession = {
  userId: string;
  actor: string;
  authType: AuthType;
  username: string | null;
  email: string | null;
  displayName: string;
  role: StaffRole;
  isOwner: boolean;
  permissions: Permission[];
  csrfToken: string;
};

export type StaffAccount = {
  id: string;
  email: string;
  displayName: string | null;
  role: StaffRole;
  enabled: boolean;
  linked: boolean;
  lastLoginAt: string | null;
  createdAt: string;
  updatedAt: string;
};

export type OAuthSettings = {
  clientId: string;
  workspaceDomain: string;
  hasClientSecret: boolean;
  callbackUrl: string;
  configured: boolean;
};

const SESSION_HOURS = 8;
const IDLE_MINUTES = 30;
const LOGIN_MAX_FAILURES = 5;
const LOGIN_LOCK_MINUTES = 15;
const GOOGLE_AUTHORIZE_URL = "https://accounts.google.com/o/oauth2/v2/auth";
const GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token";
const GOOGLE_JWKS = createRemoteJWKSet(new URL("https://www.googleapis.com/oauth2/v3/certs"));

const permissionsByRole: Record<StaffRole, Permission[]> = {
  viewer: ["view"],
  operator: ["view", "operate"],
  admin: ["view", "operate", "configure"],
};

const authSchema = [
  `CREATE TABLE IF NOT EXISTS auth_users (
    id TEXT PRIMARY KEY,
    provider TEXT NOT NULL CHECK (provider IN ('local', 'google')),
    username TEXT UNIQUE,
    email TEXT UNIQUE,
    google_sub TEXT UNIQUE,
    display_name TEXT,
    role TEXT NOT NULL CHECK (role IN ('admin', 'operator', 'viewer')),
    is_owner INTEGER NOT NULL DEFAULT 0,
    enabled INTEGER NOT NULL DEFAULT 1,
    password_hash TEXT,
    failed_attempts INTEGER NOT NULL DEFAULT 0,
    locked_until TEXT,
    last_login_at TEXT,
    created_by TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS auth_sessions (
    token_hash TEXT PRIMARY KEY,
    user_id TEXT NOT NULL REFERENCES auth_users(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL,
    absolute_expires_at TEXT NOT NULL
  )`,
  `CREATE INDEX IF NOT EXISTS auth_sessions_user_idx ON auth_sessions (user_id)`,
  `CREATE TABLE IF NOT EXISTS oauth_states (
    state_hash TEXT PRIMARY KEY,
    code_verifier TEXT NOT NULL,
    nonce TEXT NOT NULL,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS auth_settings (
    id INTEGER PRIMARY KEY DEFAULT 1,
    google_client_id TEXT,
    google_client_secret_encrypted TEXT,
    google_workspace_domain TEXT,
    updated_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS auth_rate_limits (
    bucket TEXT PRIMARY KEY,
    window_started_at TEXT NOT NULL,
    attempts INTEGER NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    occurred_at TEXT NOT NULL,
    actor TEXT NOT NULL,
    action TEXT NOT NULL,
    detail TEXT
  )`,
  `CREATE INDEX IF NOT EXISTS audit_log_occurred_idx ON audit_log (occurred_at)`,
];

let authInitialized = false;

export function ensureAuthSchema(): void {
  if (authInitialized) return;
  const database = db();
  database.transaction(() => {
    for (const statement of authSchema) database.prepare(statement).run();
  });
  database
    .prepare("INSERT OR IGNORE INTO auth_settings (id, updated_at) VALUES (1, ?)")
    .bind(new Date().toISOString())
    .run();
  database.prepare("DELETE FROM auth_sessions WHERE absolute_expires_at <= ?").bind(new Date().toISOString()).run();
  database.prepare("DELETE FROM oauth_states WHERE expires_at <= ?").bind(new Date().toISOString()).run();
  database.prepare("DELETE FROM audit_log WHERE occurred_at < ?").bind(new Date(Date.now() - 365 * 24 * 60 * 60 * 1000).toISOString()).run();
  authInitialized = true;
}

function sha256(value: string): string {
  return createHash("sha256").update(value).digest("hex");
}

function base64url(bytes: Uint8Array): string {
  return Buffer.from(bytes).toString("base64url");
}

function safeEqual(left: string, right: string): boolean {
  const a = Buffer.from(left);
  const b = Buffer.from(right);
  return a.length === b.length && timingSafeEqual(a, b);
}

function sessionSecret(): string {
  const value = process.env.SESSION_SECRET?.trim();
  if (!value || value.length < 32) {
    throw new HttpError(503, "Authentication is not configured", "SESSION_SECRET must contain at least 32 random characters.");
  }
  return value;
}

export function applicationOrigin(): string {
  const configured = process.env.APP_ORIGIN?.trim();
  if (!configured) {
    if (process.env.NODE_ENV === "production") {
      throw new HttpError(503, "Authentication is not configured", "APP_ORIGIN is required in production.");
    }
    return "http://127.0.0.1:3000";
  }

  let origin: URL;
  try {
    origin = new URL(configured);
  } catch {
    throw new HttpError(503, "APP_ORIGIN is not a valid absolute URL");
  }
  if (origin.origin !== configured.replace(/\/$/, "")) {
    throw new HttpError(503, "APP_ORIGIN must contain only a scheme and host");
  }
  if (process.env.NODE_ENV === "production" && origin.protocol !== "https:" && process.env.ALLOW_INSECURE_HTTP !== "true") {
    throw new HttpError(503, "HTTPS is required", "Use HTTPS at the internal reverse proxy, or set ALLOW_INSECURE_HTTP=true only for isolated testing.");
  }
  return origin.origin;
}

function assertExpectedHost(request: Request): void {
  const expected = new URL(applicationOrigin());
  const host = request.headers.get("host");
  if (!host || host.toLowerCase() !== expected.host.toLowerCase()) {
    throw new HttpError(403, "Unexpected request host");
  }
}

export function assertRequestOrigin(request: Request): void {
  assertExpectedHost(request);
  if (request.headers.get("sec-fetch-site") === "cross-site") {
    throw new HttpError(403, "Cross-site request blocked");
  }
  const origin = request.headers.get("origin");
  if (!origin || origin !== applicationOrigin()) {
    throw new HttpError(403, "Request origin could not be verified");
  }
}

function isUnsafeMethod(method: string): boolean {
  return !["GET", "HEAD", "OPTIONS"].includes(method.toUpperCase());
}

function cookieName(): string {
  return applicationOrigin().startsWith("https://") ? "__Host-relay_session" : "relay_session";
}

function oauthStateCookieName(): string {
  return applicationOrigin().startsWith("https://") ? "__Host-relay_oauth_state" : "relay_oauth_state";
}

function cookieValue(request: Request, name: string): string | null {
  const cookie = request.headers.get("cookie");
  if (!cookie) return null;
  for (const part of cookie.split(";")) {
    const [key, ...value] = part.trim().split("=");
    if (key === name) return decodeURIComponent(value.join("="));
  }
  return null;
}

function cookieHeader(name: string, value: string, maxAge: number): string {
  const secure = applicationOrigin().startsWith("https://") ? "; Secure" : "";
  return `${name}=${encodeURIComponent(value)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${maxAge}${secure}`;
}

export function clearSessionCookie(): string {
  return cookieHeader(cookieName(), "", 0);
}

function csrfForToken(token: string): string {
  return createHmac("sha256", sessionSecret()).update(`relay-csrf:${token}`).digest("base64url");
}

type UserRow = {
  id: string;
  provider: AuthType;
  username: string | null;
  email: string | null;
  google_sub: string | null;
  display_name: string | null;
  role: StaffRole;
  is_owner: number;
  enabled: number;
  password_hash: string | null;
  failed_attempts: number;
  locked_until: string | null;
  last_login_at: string | null;
  created_at: string;
  updated_at: string;
};

type SessionRow = UserRow & {
  token_hash: string;
  last_seen_at: string;
  absolute_expires_at: string;
};

function sessionFromRow(row: SessionRow, rawToken: string): AuthSession {
  const isOwner = Boolean(row.is_owner);
  const permissions = isOwner
    ? (["view", "operate", "configure", "manage_access"] satisfies Permission[])
    : permissionsByRole[row.role];
  return {
    userId: row.id,
    actor: row.provider === "local" ? `local:${row.username}` : `google:${row.email}`,
    authType: row.provider,
    username: row.username,
    email: row.email,
    displayName: row.display_name || row.username || row.email || "Administrator",
    role: row.role,
    isOwner,
    permissions,
    csrfToken: csrfForToken(rawToken),
  };
}

function createSession(userId: string): { session: AuthSession; cookie: string } {
  ensureAuthSchema();
  const token = base64url(randomBytes(32));
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_HOURS * 60 * 60 * 1000);
  db()
    .prepare("INSERT INTO auth_sessions (token_hash, user_id, created_at, last_seen_at, absolute_expires_at) VALUES (?, ?, ?, ?, ?)")
    .bind(sha256(token), userId, now.toISOString(), now.toISOString(), expires.toISOString())
    .run();
  const row = getSessionRow(sha256(token));
  if (!row) throw new HttpError(500, "The administrator session could not be created");
  return {
    session: sessionFromRow(row, token),
    cookie: cookieHeader(cookieName(), token, SESSION_HOURS * 60 * 60),
  };
}

function getSessionRow(tokenHash: string): SessionRow | null {
  return db()
    .prepare(`SELECT s.token_hash, s.last_seen_at, s.absolute_expires_at,
      u.id, u.provider, u.username, u.email, u.google_sub, u.display_name, u.role, u.is_owner,
      u.enabled, u.password_hash, u.failed_attempts, u.locked_until, u.last_login_at, u.created_at, u.updated_at
      FROM auth_sessions s JOIN auth_users u ON u.id = s.user_id WHERE s.token_hash = ?`)
    .bind(tokenHash)
    .first<SessionRow>();
}

export function currentSession(request: Request): AuthSession | null {
  ensureAuthSchema();
  assertExpectedHost(request);
  const token = cookieValue(request, cookieName());
  if (!token) return null;
  const tokenHash = sha256(token);
  const row = getSessionRow(tokenHash);
  if (!row) return null;

  const now = Date.now();
  const idleExpiry = Date.parse(row.last_seen_at) + IDLE_MINUTES * 60 * 1000;
  if (!row.enabled || Date.parse(row.absolute_expires_at) <= now || idleExpiry <= now) {
    db().prepare("DELETE FROM auth_sessions WHERE token_hash = ?").bind(tokenHash).run();
    return null;
  }

  if (now - Date.parse(row.last_seen_at) > 5 * 60 * 1000) {
    db().prepare("UPDATE auth_sessions SET last_seen_at = ? WHERE token_hash = ?").bind(new Date(now).toISOString(), tokenHash).run();
  }
  return sessionFromRow(row, token);
}

export function requireSession(request: Request, permission: Permission = "view"): AuthSession {
  const session = currentSession(request);
  if (!session) throw new HttpError(401, "Sign in required");
  if (!session.permissions.includes(permission)) throw new HttpError(403, "You do not have permission to perform this action");

  if (isUnsafeMethod(request.method)) {
    assertRequestOrigin(request);
    const supplied = request.headers.get("x-csrf-token") ?? "";
    if (!supplied || !safeEqual(supplied, session.csrfToken)) {
      throw new HttpError(403, "The security token is missing or invalid");
    }
  }
  return session;
}

export async function requestActor(request: Request, permission: Permission = "view"): Promise<string> {
  return requireSession(request, permission).actor;
}

export function requestScheduler(request: Request): string {
  const configured = process.env.SCHEDULER_TOKEN?.trim();
  if (!configured || configured.length < 32) {
    throw new HttpError(503, "Scheduler authentication is not configured");
  }
  const header = request.headers.get("authorization") ?? "";
  const supplied = header.startsWith("Bearer ") ? header.slice(7) : "";
  if (!supplied || !safeEqual(supplied, configured)) throw new HttpError(401, "Scheduler authentication failed");
  return "scheduler";
}

function verifyPassword(password: string, encoded: string): boolean {
  const [algorithm, iterationsValue, salt, expected] = encoded.split("$");
  const iterations = Number(iterationsValue);
  if (algorithm !== "pbkdf2-sha256" || !Number.isInteger(iterations) || iterations < 100_000 || !salt || !expected) {
    throw new HttpError(503, "The local administrator password hash is invalid");
  }
  const actual = pbkdf2Sync(password, Buffer.from(salt, "base64url"), iterations, 32, "sha256").toString("base64url");
  return safeEqual(actual, expected);
}

function addAuthAudit(actor: string, action: string, detail?: string): void {
  ensureAuthSchema();
  db()
    .prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, ?, ?)")
    .bind(new Date().toISOString(), actor, action, detail ?? null)
    .run();
}

function consumeRateLimit(bucket: string, maximum: number, windowSeconds: number): void {
  ensureAuthSchema();
  const database = db();
  const now = Date.now();
  database.transaction(() => {
    const row = database.prepare("SELECT window_started_at, attempts FROM auth_rate_limits WHERE bucket = ?")
      .bind(bucket).first<{ window_started_at: string; attempts: number }>();
    if (!row || Date.parse(row.window_started_at) + windowSeconds * 1000 <= now) {
      database.prepare(`INSERT INTO auth_rate_limits (bucket, window_started_at, attempts) VALUES (?, ?, 1)
        ON CONFLICT(bucket) DO UPDATE SET window_started_at = excluded.window_started_at, attempts = 1`)
        .bind(bucket, new Date(now).toISOString()).run();
      return;
    }
    if (row.attempts >= maximum) throw new HttpError(429, "Too many authentication attempts", "Wait briefly before trying again.");
    database.prepare("UPDATE auth_rate_limits SET attempts = attempts + 1 WHERE bucket = ?").bind(bucket).run();
  });
}

export function localLogin(username: string, password: string): { session: AuthSession; cookie: string } {
  ensureAuthSchema();
  consumeRateLimit("local-login-global", 20, 60);
  const normalizedUsername = username.trim().toLowerCase();
  const row = db()
    .prepare("SELECT * FROM auth_users WHERE provider = 'local' AND lower(username) = ? LIMIT 1")
    .bind(normalizedUsername)
    .first<UserRow>();

  if (!row || !row.password_hash || !row.enabled) {
    pbkdf2Sync(password, Buffer.from("relay-invalid-user"), 100_000, 32, "sha256");
    addAuthAudit("anonymous", "authentication.failed", "Invalid local administrator credentials");
    throw new HttpError(401, "Invalid username or password");
  }

  if (row.locked_until && Date.parse(row.locked_until) > Date.now()) {
    throw new HttpError(429, "The local administrator account is temporarily locked", "Wait 15 minutes before trying again.");
  }

  if (!verifyPassword(password, row.password_hash)) {
    const failures = row.failed_attempts + 1;
    const lockedUntil = failures >= LOGIN_MAX_FAILURES
      ? new Date(Date.now() + LOGIN_LOCK_MINUTES * 60 * 1000).toISOString()
      : null;
    db()
      .prepare("UPDATE auth_users SET failed_attempts = ?, locked_until = ?, updated_at = ? WHERE id = ?")
      .bind(failures, lockedUntil, new Date().toISOString(), row.id)
      .run();
    addAuthAudit(`local:${row.username}`, "authentication.failed", "Invalid local administrator credentials");
    throw new HttpError(401, "Invalid username or password");
  }

  const now = new Date().toISOString();
  db()
    .prepare("UPDATE auth_users SET failed_attempts = 0, locked_until = NULL, last_login_at = ?, updated_at = ? WHERE id = ?")
    .bind(now, now, row.id)
    .run();
  db().prepare("DELETE FROM auth_sessions WHERE user_id = ?").bind(row.id).run();
  const result = createSession(row.id);
  addAuthAudit(result.session.actor, "authentication.succeeded", "Local administrator signed in");
  return result;
}

export function logout(request: Request): void {
  const session = requireSession(request, "view");
  const token = cookieValue(request, cookieName());
  if (token) db().prepare("DELETE FROM auth_sessions WHERE token_hash = ?").bind(sha256(token)).run();
  addAuthAudit(session.actor, "authentication.logout");
}

export function changeLocalPassword(actor: AuthSession, currentPassword: string, nextPassword: string): void {
  if (!actor.isOwner || actor.authType !== "local") throw new HttpError(403, "Only the local administrator can change this password");
  if (nextPassword.length < 14) throw new HttpError(400, "The new password must be at least 14 characters");
  const row = db().prepare("SELECT * FROM auth_users WHERE id = ?").bind(actor.userId).first<UserRow>();
  if (!row?.password_hash || !verifyPassword(currentPassword, row.password_hash)) {
    throw new HttpError(401, "The current password is incorrect");
  }
  const salt = randomBytes(16);
  const iterations = 600_000;
  const derived = pbkdf2Sync(nextPassword, salt, iterations, 32, "sha256").toString("base64url");
  const encoded = `pbkdf2-sha256$${iterations}$${salt.toString("base64url")}$${derived}`;
  db().transaction(() => {
    db().prepare("UPDATE auth_users SET password_hash = ?, updated_at = ? WHERE id = ?").bind(encoded, new Date().toISOString(), actor.userId).run();
    db().prepare("DELETE FROM auth_sessions WHERE user_id = ?").bind(actor.userId).run();
  });
  addAuthAudit(actor.actor, "authentication.password_changed");
}

type OAuthSettingsRow = {
  google_client_id: string | null;
  google_client_secret_encrypted: string | null;
  google_workspace_domain: string | null;
};

async function oauthSettingsWithSecret(): Promise<OAuthSettings & { clientSecret: string }> {
  ensureAuthSchema();
  const row = db().prepare("SELECT * FROM auth_settings WHERE id = 1").first<OAuthSettingsRow>();
  const clientId = process.env.GOOGLE_OAUTH_CLIENT_ID?.trim() || row?.google_client_id?.trim() || "";
  const workspaceDomain = process.env.GOOGLE_OAUTH_DOMAIN?.trim().toLowerCase() || row?.google_workspace_domain?.trim().toLowerCase() || "";
  const clientSecret = process.env.GOOGLE_OAUTH_CLIENT_SECRET?.trim()
    || (row?.google_client_secret_encrypted ? await decryptSecret(row.google_client_secret_encrypted) : "");
  return {
    clientId,
    clientSecret,
    workspaceDomain,
    hasClientSecret: Boolean(clientSecret),
    callbackUrl: `${applicationOrigin()}/api/auth/google/callback`,
    configured: Boolean(clientId && clientSecret && workspaceDomain),
  };
}

export async function getOAuthSettings(): Promise<OAuthSettings> {
  const settings = await oauthSettingsWithSecret();
  return {
    clientId: settings.clientId,
    workspaceDomain: settings.workspaceDomain,
    hasClientSecret: settings.hasClientSecret,
    callbackUrl: settings.callbackUrl,
    configured: settings.configured,
  };
}

export async function saveOAuthSettings(input: { clientId?: string; clientSecret?: string; workspaceDomain?: string }, actor: string): Promise<OAuthSettings> {
  ensureAuthSchema();
  if (process.env.GOOGLE_OAUTH_CLIENT_ID || process.env.GOOGLE_OAUTH_CLIENT_SECRET || process.env.GOOGLE_OAUTH_DOMAIN) {
    throw new HttpError(409, "Google sign-in settings are controlled by environment variables");
  }
  const current = await oauthSettingsWithSecret();
  const clientId = (input.clientId ?? current.clientId).trim();
  const domain = (input.workspaceDomain ?? current.workspaceDomain).trim().toLowerCase();
  if (clientId && !clientId.endsWith(".apps.googleusercontent.com")) throw new HttpError(400, "Enter a valid Google web OAuth client ID");
  if (domain && (!domain.includes(".") || domain.includes("@"))) throw new HttpError(400, "Enter the Google Workspace domain, such as school.edu.au");
  const encrypted = input.clientSecret?.trim() ? await encryptSecret(input.clientSecret.trim()) : null;
  db()
    .prepare(`UPDATE auth_settings SET google_client_id = ?, google_client_secret_encrypted = COALESCE(?, google_client_secret_encrypted),
      google_workspace_domain = ?, updated_at = ? WHERE id = 1`)
    .bind(clientId || null, encrypted, domain || null, new Date().toISOString())
    .run();
  addAuthAudit(actor, "access.oauth_settings_updated", "Google Workspace sign-in settings were updated");
  return getOAuthSettings();
}

export async function beginGoogleOAuth(request: Request): Promise<{ url: string; cookie: string }> {
  ensureAuthSchema();
  consumeRateLimit("google-oauth-start-global", 60, 60);
  assertExpectedHost(request);
  if (request.headers.get("sec-fetch-site") === "cross-site") throw new HttpError(403, "Cross-site request blocked");
  const settings = await oauthSettingsWithSecret();
  if (!settings.configured) throw new HttpError(503, "Google Workspace sign-in has not been configured");
  const state = base64url(randomBytes(32));
  const verifier = base64url(randomBytes(48));
  const nonce = base64url(randomBytes(32));
  const challenge = createHash("sha256").update(verifier).digest("base64url");
  const now = new Date();
  const expires = new Date(now.getTime() + 10 * 60 * 1000);
  db().prepare("DELETE FROM oauth_states WHERE expires_at <= ?").bind(now.toISOString()).run();
  db()
    .prepare("INSERT INTO oauth_states (state_hash, code_verifier, nonce, created_at, expires_at) VALUES (?, ?, ?, ?, ?)")
    .bind(sha256(state), verifier, nonce, now.toISOString(), expires.toISOString())
    .run();

  const url = new URL(GOOGLE_AUTHORIZE_URL);
  url.search = new URLSearchParams({
    client_id: settings.clientId,
    redirect_uri: settings.callbackUrl,
    response_type: "code",
    scope: "openid email profile",
    state,
    nonce,
    code_challenge: challenge,
    code_challenge_method: "S256",
    access_type: "online",
    prompt: "select_account",
    hd: settings.workspaceDomain,
  }).toString();
  return { url: url.toString(), cookie: cookieHeader(oauthStateCookieName(), state, 10 * 60) };
}

type OAuthStateRow = { code_verifier: string; nonce: string; expires_at: string };

export async function completeGoogleOAuth(request: Request): Promise<{ session: AuthSession; cookie: string }> {
  ensureAuthSchema();
  assertExpectedHost(request);
  const url = new URL(request.url);
  const state = url.searchParams.get("state") ?? "";
  const code = url.searchParams.get("code") ?? "";
  const stateCookie = cookieValue(request, oauthStateCookieName()) ?? "";
  if (!state || !code || !stateCookie || !safeEqual(state, stateCookie)) {
    throw new HttpError(400, "Google sign-in state could not be verified");
  }
  const stateHash = sha256(state);
  const stored = db().prepare("SELECT code_verifier, nonce, expires_at FROM oauth_states WHERE state_hash = ?").bind(stateHash).first<OAuthStateRow>();
  db().prepare("DELETE FROM oauth_states WHERE state_hash = ?").bind(stateHash).run();
  if (!stored || Date.parse(stored.expires_at) <= Date.now()) throw new HttpError(400, "Google sign-in has expired; start again");

  const settings = await oauthSettingsWithSecret();
  const response = await fetch(GOOGLE_TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      code,
      client_id: settings.clientId,
      client_secret: settings.clientSecret,
      redirect_uri: settings.callbackUrl,
      grant_type: "authorization_code",
      code_verifier: stored.code_verifier,
    }),
  });
  const token = await response.json() as { id_token?: string; error?: string; error_description?: string };
  if (!response.ok || !token.id_token) throw new HttpError(401, "Google sign-in failed", token.error_description || token.error);

  const { payload } = await jwtVerify(token.id_token, GOOGLE_JWKS, {
    issuer: ["https://accounts.google.com", "accounts.google.com"],
    audience: settings.clientId,
  });
  const email = typeof payload.email === "string" ? payload.email.trim().toLowerCase() : "";
  const subject = typeof payload.sub === "string" ? payload.sub : "";
  const domain = typeof payload.hd === "string" ? payload.hd.trim().toLowerCase() : "";
  if (payload.nonce !== stored.nonce || payload.email_verified !== true || !email || !subject || domain !== settings.workspaceDomain) {
    throw new HttpError(403, "The Google Workspace identity could not be verified");
  }

  let user = db().prepare("SELECT * FROM auth_users WHERE provider = 'google' AND google_sub = ? LIMIT 1").bind(subject).first<UserRow>();
  if (user && user.email?.trim().toLowerCase() !== email) {
    addAuthAudit(`google:${email}`, "authentication.denied", "Google subject no longer matches its allowlisted email address");
    throw new HttpError(
      403,
      "Access has not been granted for this Google account",
      "The account identity changed after it was approved. Ask the local Relay administrator to review the IT access entry.",
    );
  }
  if (!user) {
    user = db().prepare("SELECT * FROM auth_users WHERE provider = 'google' AND lower(email) = ? LIMIT 1").bind(email).first<UserRow>();
    if (user && user.google_sub && user.google_sub !== subject) user = null;
    if (user && !user.google_sub) {
      db().prepare("UPDATE auth_users SET google_sub = ?, updated_at = ? WHERE id = ?").bind(subject, new Date().toISOString(), user.id).run();
      user.google_sub = subject;
    }
  }
  if (!user || !user.enabled) {
    addAuthAudit(`google:${email}`, "authentication.denied", "Google account is not enabled for Relay");
    throw new HttpError(403, "Access has not been granted for this Google account", "Ask the local Relay administrator to add and enable your email address.");
  }

  const now = new Date().toISOString();
  const displayName = typeof payload.name === "string" && payload.name.trim() ? payload.name.trim() : user.display_name;
  db()
    .prepare("UPDATE auth_users SET display_name = COALESCE(?, display_name), last_login_at = ?, updated_at = ? WHERE id = ?")
    .bind(displayName || null, now, now, user.id)
    .run();
  db().prepare("DELETE FROM auth_sessions WHERE user_id = ?").bind(user.id).run();
  const result = createSession(user.id);
  addAuthAudit(result.session.actor, "authentication.succeeded", "Google Workspace administrator signed in");
  return result;
}

export function listStaffAccounts(): StaffAccount[] {
  ensureAuthSchema();
  const rows = db()
    .prepare(`SELECT id, email, display_name, role, enabled, google_sub, last_login_at, created_at, updated_at
      FROM auth_users WHERE provider = 'google' ORDER BY lower(email)`)
    .all<{
      id: string; email: string; display_name: string | null; role: StaffRole; enabled: number;
      google_sub: string | null; last_login_at: string | null; created_at: string; updated_at: string;
    }>().results;
  return rows.map((row) => ({
    id: row.id,
    email: row.email,
    displayName: row.display_name,
    role: row.role,
    enabled: Boolean(row.enabled),
    linked: Boolean(row.google_sub),
    lastLoginAt: row.last_login_at,
    createdAt: row.created_at,
    updatedAt: row.updated_at,
  }));
}

function normalizeRole(value: string): StaffRole {
  if (value === "admin" || value === "operator" || value === "viewer") return value;
  throw new HttpError(400, "Choose Administrator, Operator, or Viewer");
}

export function saveStaffAccount(input: { id?: string; email?: string; displayName?: string; role?: string; enabled?: boolean }, actor: string): StaffAccount {
  ensureAuthSchema();
  const now = new Date().toISOString();
  const existing = input.id
    ? db().prepare("SELECT * FROM auth_users WHERE id = ? AND provider = 'google'").bind(input.id).first<UserRow>()
    : null;
  const email = (input.email ?? existing?.email ?? "").trim().toLowerCase();
  const displayName = (input.displayName ?? existing?.display_name ?? "").trim();
  const role = normalizeRole(input.role ?? existing?.role ?? "viewer");
  const enabled = input.enabled ?? (existing ? Boolean(existing.enabled) : true);
  if (!/^\S+@\S+\.\S+$/.test(email)) throw new HttpError(400, "Enter a valid Google Workspace email address");

  let id: string;
  try {
    if (existing) {
      id = existing.id;
      db().transaction(() => {
        db().prepare(`UPDATE auth_users SET email = ?, display_name = ?, role = ?, enabled = ?,
          google_sub = CASE WHEN lower(email) = ? THEN google_sub ELSE NULL END, updated_at = ? WHERE id = ?`)
          .bind(email, displayName || null, role, Number(enabled), email, now, id).run();
        db().prepare("DELETE FROM auth_sessions WHERE user_id = ?").bind(id).run();
      });
    } else {
      id = crypto.randomUUID();
      db().prepare(`INSERT INTO auth_users
        (id, provider, email, display_name, role, is_owner, enabled, created_by, created_at, updated_at)
        VALUES (?, 'google', ?, ?, ?, 0, ?, ?, ?, ?)`)
        .bind(id, email, displayName || null, role, Number(enabled), actor, now, now).run();
    }
  } catch (error) {
    if (error instanceof Error && error.message.includes("UNIQUE")) throw new HttpError(409, "That Google account is already configured");
    throw error;
  }
  addAuthAudit(actor, existing ? "access.staff_updated" : "access.staff_added", `${email} (${role}, ${enabled ? "enabled" : "disabled"})`);
  return listStaffAccounts().find((account) => account.id === id)!;
}

export function deleteStaffAccount(id: string, actor: string): void {
  ensureAuthSchema();
  const account = db().prepare("SELECT email FROM auth_users WHERE id = ? AND provider = 'google'").bind(id).first<{ email: string }>();
  if (!account) throw new HttpError(404, "Staff account not found");
  db().prepare("DELETE FROM auth_users WHERE id = ?").bind(id).run();
  addAuthAudit(actor, "access.staff_removed", account.email);
}

export function authReadiness(): { localAdministrator: boolean; googleSignInConfigured: boolean } {
  ensureAuthSchema();
  const owner = db().prepare("SELECT id FROM auth_users WHERE provider = 'local' AND is_owner = 1 AND enabled = 1 LIMIT 1").first<{ id: string }>();
  const settings = db().prepare("SELECT google_client_id, google_client_secret_encrypted, google_workspace_domain FROM auth_settings WHERE id = 1").first<OAuthSettingsRow>();
  const environmentConfigured = Boolean(process.env.GOOGLE_OAUTH_CLIENT_ID && process.env.GOOGLE_OAUTH_CLIENT_SECRET && process.env.GOOGLE_OAUTH_DOMAIN);
  return {
    localAdministrator: Boolean(owner),
    googleSignInConfigured: environmentConfigured || Boolean(settings?.google_client_id && settings.google_client_secret_encrypted && settings.google_workspace_domain),
  };
}

export function deploymentReady(): boolean {
  try {
    applicationOrigin();
    sessionSecret();
    const scheduler = process.env.SCHEDULER_TOKEN?.trim() ?? "";
    const encryption = process.env.CONFIG_ENCRYPTION_KEY?.trim() ?? "";
    if (scheduler.length < 32 || Buffer.from(encryption, "base64").byteLength !== 32) return false;
    return authReadiness().localAdministrator;
  } catch {
    return false;
  }
}
