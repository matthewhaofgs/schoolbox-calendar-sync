import { db } from "./db";
import { decryptSecret, encryptSecret, HttpError } from "./security";

export type AppConfig = {
  schoolboxBaseUrl: string;
  schoolboxToken?: string;
  googleServiceAccountJson?: string;
  googleAdminEmail: string;
  googleCustomer: string;
  timezone: string;
  pastDays: number;
  futureDays: number;
  concurrency: number;
  syncIntervalMinutes: number;
  syncNewUsersByDefault: boolean;
  enabled: boolean;
  setupCompleted: boolean;
  hasSchoolboxToken: boolean;
  hasGoogleServiceAccount: boolean;
  serviceAccountEmail?: string;
  serviceAccountClientId?: string;
  updatedAt?: string;
};

export type ConfigInput = Partial<Omit<AppConfig, "hasSchoolboxToken" | "hasGoogleServiceAccount" | "serviceAccountEmail" | "serviceAccountClientId">>;

export type RunSummary = {
  id: string;
  trigger: string;
  status: string;
  startedAt: string;
  completedAt: string | null;
  usersDiscovered: number;
  usersMatched: number;
  usersSynced: number;
  eventsCreated: number;
  eventsUpdated: number;
  eventsDeleted: number;
  eventsUnchanged: number;
  errors: number;
  message: string | null;
};

export type UserMapping = {
  googleUserId: string;
  googleEmail: string;
  schoolboxUserId: number | null;
  schoolboxEmail: string | null;
  displayName: string | null;
  role: string | null;
  status: string;
  lastSyncAt: string | null;
  lastError: string | null;
  eventCount: number;
  syncEnabled: boolean;
  directoryActive: boolean;
  updatedAt: string;
};

export type EventMapping = {
  googleUserId: string;
  sourceKey: string;
  googleEventId: string;
  sourceHash: string;
  sourceStart: string;
  sourceEnd: string;
  lastSeenRunId: string;
  createdAt: string;
  updatedAt: string;
};

const schemaStatements = [
  `CREATE TABLE IF NOT EXISTS app_config (
    id INTEGER PRIMARY KEY DEFAULT 1,
    schoolbox_base_url TEXT,
    schoolbox_token_encrypted TEXT,
    google_service_account_encrypted TEXT,
    google_admin_email TEXT,
    google_customer TEXT NOT NULL DEFAULT 'my_customer',
    timezone TEXT NOT NULL DEFAULT 'Australia/Sydney',
    past_days INTEGER NOT NULL DEFAULT 30,
    future_days INTEGER NOT NULL DEFAULT 180,
    concurrency INTEGER NOT NULL DEFAULT 3,
    sync_interval_minutes INTEGER NOT NULL DEFAULT 360,
    sync_new_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_users_by_default IN (0, 1)),
    enabled INTEGER NOT NULL DEFAULT 0,
    setup_completed INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS sync_runs (
    id TEXT PRIMARY KEY,
    trigger TEXT NOT NULL,
    status TEXT NOT NULL,
    started_at TEXT NOT NULL,
    heartbeat_at TEXT,
    completed_at TEXT,
    users_discovered INTEGER NOT NULL DEFAULT 0,
    users_matched INTEGER NOT NULL DEFAULT 0,
    users_synced INTEGER NOT NULL DEFAULT 0,
    events_created INTEGER NOT NULL DEFAULT 0,
    events_updated INTEGER NOT NULL DEFAULT 0,
    events_deleted INTEGER NOT NULL DEFAULT 0,
    events_unchanged INTEGER NOT NULL DEFAULT 0,
    errors INTEGER NOT NULL DEFAULT 0,
    message TEXT
  )`,
  `CREATE TABLE IF NOT EXISTS user_mappings (
    google_user_id TEXT PRIMARY KEY,
    google_email TEXT NOT NULL,
    schoolbox_user_id INTEGER,
    schoolbox_email TEXT,
    display_name TEXT,
    role TEXT,
    status TEXT NOT NULL DEFAULT 'unmatched',
    last_sync_at TEXT,
    last_error TEXT,
    event_count INTEGER NOT NULL DEFAULT 0,
    sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK (sync_enabled IN (0, 1)),
    directory_active INTEGER NOT NULL DEFAULT 1 CHECK (directory_active IN (0, 1)),
    updated_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS event_mappings (
    google_user_id TEXT NOT NULL,
    source_key TEXT NOT NULL,
    google_event_id TEXT NOT NULL,
    source_hash TEXT NOT NULL,
    source_start TEXT NOT NULL,
    source_end TEXT NOT NULL,
    last_seen_run_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (google_user_id, source_key)
  )`,
  `CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    occurred_at TEXT NOT NULL,
    actor TEXT NOT NULL,
    action TEXT NOT NULL,
    detail TEXT
  )`,
  "CREATE INDEX IF NOT EXISTS event_mappings_seen_idx ON event_mappings (google_user_id, last_seen_run_id)",
  "CREATE INDEX IF NOT EXISTS sync_runs_started_idx ON sync_runs (started_at DESC)",
];

let initialized = false;

export async function ensureSchema(): Promise<void> {
  if (initialized) return;
  const binding = db();
  binding.transaction(() => {
    for (const statement of schemaStatements) binding.prepare(statement).run();
  });
  const runColumns = binding.prepare("PRAGMA table_info(sync_runs)").all<{ name: string }>().results;
  if (!runColumns.some((column) => column.name === "heartbeat_at")) {
    binding.prepare("ALTER TABLE sync_runs ADD COLUMN heartbeat_at TEXT").run();
  }
  const configColumns = binding.prepare("PRAGMA table_info(app_config)").all<{ name: string }>().results;
  if (!configColumns.some((column) => column.name === "sync_new_users_by_default")) {
    // Legacy installations implicitly synced every newly discovered account. Keep
    // that behaviour on upgrade; brand-new databases use the safer CREATE default.
    binding.prepare("ALTER TABLE app_config ADD COLUMN sync_new_users_by_default INTEGER NOT NULL DEFAULT 1 CHECK (sync_new_users_by_default IN (0, 1))").run();
  }
  let userColumns = binding.prepare("PRAGMA table_info(user_mappings)").all<{ name: string }>().results;
  const emailHasGlobalUniqueIndex = binding.prepare("PRAGMA index_list(user_mappings)")
    .all<{ name: string; unique: number; partial: number }>().results
    .filter((index) => index.unique && !index.partial)
    .some((index) => {
      const escapedName = index.name.replaceAll('"', '""');
      const columns = binding.prepare(`PRAGMA index_info("${escapedName}")`).all<{ name: string }>().results;
      return columns.length === 1 && columns[0]?.name === "google_email";
    });
  if (!userColumns.some((column) => column.name === "directory_active") || emailHasGlobalUniqueIndex) {
    const hasSyncEnabled = userColumns.some((column) => column.name === "sync_enabled");
    const hasDirectoryActive = userColumns.some((column) => column.name === "directory_active");
    binding.transaction(() => {
      binding.prepare("DROP TABLE IF EXISTS user_mappings_rebuilt").run();
      binding.prepare(`CREATE TABLE user_mappings_rebuilt (
        google_user_id TEXT PRIMARY KEY,
        google_email TEXT NOT NULL,
        schoolbox_user_id INTEGER,
        schoolbox_email TEXT,
        display_name TEXT,
        role TEXT,
        status TEXT NOT NULL DEFAULT 'unmatched',
        last_sync_at TEXT,
        last_error TEXT,
        event_count INTEGER NOT NULL DEFAULT 0,
        sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK (sync_enabled IN (0, 1)),
        directory_active INTEGER NOT NULL DEFAULT 1 CHECK (directory_active IN (0, 1)),
        updated_at TEXT NOT NULL
      )`).run();
      binding.prepare(`INSERT INTO user_mappings_rebuilt
        (google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name, role, status,
         last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
        SELECT google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name, role, status,
         last_sync_at, last_error, event_count, ${hasSyncEnabled ? "sync_enabled" : "1"},
         ${hasDirectoryActive ? "directory_active" : "1"}, updated_at
        FROM user_mappings`).run();
      binding.prepare("DROP TABLE user_mappings").run();
      binding.prepare("ALTER TABLE user_mappings_rebuilt RENAME TO user_mappings").run();
    });
    userColumns = binding.prepare("PRAGMA table_info(user_mappings)").all<{ name: string }>().results;
  }
  if (!userColumns.some((column) => column.name === "sync_enabled")) {
    // Every previously discovered user was eligible before this policy existed.
    binding.prepare("ALTER TABLE user_mappings ADD COLUMN sync_enabled INTEGER NOT NULL DEFAULT 1 CHECK (sync_enabled IN (0, 1))").run();
  }
  binding.prepare(`CREATE UNIQUE INDEX IF NOT EXISTS user_mappings_active_email_idx
    ON user_mappings (google_email COLLATE NOCASE) WHERE directory_active = 1`).run();
  await binding
    .prepare("INSERT OR IGNORE INTO app_config (id, updated_at) VALUES (1, ?)")
    .bind(new Date().toISOString())
    .run();
  initialized = true;
}

type ConfigRow = {
  schoolbox_base_url: string | null;
  schoolbox_token_encrypted: string | null;
  google_service_account_encrypted: string | null;
  google_admin_email: string | null;
  google_customer: string;
  timezone: string;
  past_days: number;
  future_days: number;
  concurrency: number;
  sync_interval_minutes: number;
  sync_new_users_by_default: number;
  enabled: number;
  setup_completed: number;
  updated_at: string;
};

export async function getConfig(includeSecrets = false): Promise<AppConfig> {
  await ensureSchema();
  const row = await db().prepare("SELECT * FROM app_config WHERE id = 1").first<ConfigRow>();
  if (!row) throw new HttpError(500, "Application configuration row is missing");

  let serviceAccountEmail: string | undefined;
  let serviceAccountClientId: string | undefined;
  let googleServiceAccountJson: string | undefined;
  if (row.google_service_account_encrypted) {
    try {
      googleServiceAccountJson = await decryptSecret(row.google_service_account_encrypted);
      const parsed = JSON.parse(googleServiceAccountJson) as { client_email?: string; client_id?: string };
      serviceAccountEmail = parsed.client_email;
      serviceAccountClientId = parsed.client_id;
    } catch {
      if (includeSecrets) throw new HttpError(500, "Stored Google credential could not be decrypted");
    }
  }

  const result: AppConfig = {
    schoolboxBaseUrl: row.schoolbox_base_url ?? "",
    googleAdminEmail: row.google_admin_email ?? "",
    googleCustomer: row.google_customer,
    timezone: row.timezone,
    pastDays: row.past_days,
    futureDays: row.future_days,
    concurrency: row.concurrency,
    syncIntervalMinutes: row.sync_interval_minutes,
    syncNewUsersByDefault: Boolean(row.sync_new_users_by_default),
    enabled: Boolean(row.enabled),
    setupCompleted: Boolean(row.setup_completed),
    hasSchoolboxToken: Boolean(row.schoolbox_token_encrypted),
    hasGoogleServiceAccount: Boolean(row.google_service_account_encrypted),
    serviceAccountEmail,
    serviceAccountClientId,
    updatedAt: row.updated_at,
  };

  if (includeSecrets) {
    if (row.schoolbox_token_encrypted) result.schoolboxToken = await decryptSecret(row.schoolbox_token_encrypted);
    result.googleServiceAccountJson = googleServiceAccountJson;
  }
  return result;
}

export async function getStoredSchoolboxConnection(): Promise<{ baseUrl: string; token?: string }> {
  await ensureSchema();
  const row = await db()
    .prepare("SELECT schoolbox_base_url, schoolbox_token_encrypted FROM app_config WHERE id = 1")
    .first<{ schoolbox_base_url: string | null; schoolbox_token_encrypted: string | null }>();
  return {
    baseUrl: row?.schoolbox_base_url ?? "",
    token: row?.schoolbox_token_encrypted ? await decryptSecret(row.schoolbox_token_encrypted) : undefined,
  };
}

export async function getStoredGoogleConnection(): Promise<{ serviceAccountJson?: string; adminEmail: string }> {
  await ensureSchema();
  const row = await db()
    .prepare("SELECT google_service_account_encrypted, google_admin_email FROM app_config WHERE id = 1")
    .first<{ google_service_account_encrypted: string | null; google_admin_email: string | null }>();
  return {
    serviceAccountJson: row?.google_service_account_encrypted
      ? await decryptSecret(row.google_service_account_encrypted)
      : undefined,
    adminEmail: row?.google_admin_email ?? "",
  };
}

function clampInteger(value: number | undefined, fallback: number, min: number, max: number): number {
  if (value === undefined || !Number.isFinite(value)) return fallback;
  return Math.max(min, Math.min(max, Math.round(value)));
}

export async function saveConfig(input: ConfigInput, actor: string): Promise<AppConfig> {
  const current = await getConfig(false);
  const now = new Date().toISOString();
  const baseUrl = (input.schoolboxBaseUrl ?? current.schoolboxBaseUrl).trim().replace(/\/$/, "");
  if (baseUrl && current.schoolboxBaseUrl && !input.schoolboxToken) {
    let originChanged = false;
    try { originChanged = new URL(baseUrl).origin !== new URL(current.schoolboxBaseUrl).origin; } catch { originChanged = true; }
    if (originChanged) throw new HttpError(400, "Enter a new Schoolbox JWT when changing the Schoolbox host");
  }
  const adminEmail = (input.googleAdminEmail ?? current.googleAdminEmail).trim().toLowerCase();
  const customer = (input.googleCustomer ?? current.googleCustomer ?? "my_customer").trim();
  const timezone = (input.timezone ?? current.timezone).trim() || "Australia/Sydney";
  const tokenEncrypted = input.schoolboxToken
    ? await encryptSecret(input.schoolboxToken.trim())
    : null;
  let serviceAccountEncrypted: string | null = null;
  if (input.googleServiceAccountJson) {
    let parsed: { type?: string; client_email?: string; private_key?: string; client_id?: string };
    try {
      parsed = JSON.parse(input.googleServiceAccountJson);
    } catch {
      throw new HttpError(400, "The Google service account file is not valid JSON");
    }
    if (parsed.type !== "service_account" || !parsed.client_email || !parsed.private_key || !parsed.client_id) {
      throw new HttpError(400, "The Google credential is missing service-account fields");
    }
    serviceAccountEncrypted = await encryptSecret(JSON.stringify(parsed));
  }

  const activating = input.setupCompleted === true || input.enabled === true;
  if (activating) {
    const hasSchoolboxToken = Boolean(input.schoolboxToken?.trim() || current.hasSchoolboxToken);
    const hasGoogleCredential = Boolean(input.googleServiceAccountJson?.trim() || current.hasGoogleServiceAccount);
    if (!baseUrl || !hasSchoolboxToken || !adminEmail || !hasGoogleCredential) {
      throw new HttpError(400, "Complete both Schoolbox and Google Workspace connections before activating Relay");
    }
    let schoolboxUrl: URL;
    try { schoolboxUrl = new URL(baseUrl); } catch { throw new HttpError(400, "Enter a valid Schoolbox URL"); }
    if (schoolboxUrl.protocol !== "https:") throw new HttpError(400, "Schoolbox must use HTTPS");
  }

  await ensureSchema();
  await db()
    .prepare(`UPDATE app_config SET
      schoolbox_base_url = ?,
      schoolbox_token_encrypted = COALESCE(?, schoolbox_token_encrypted),
      google_service_account_encrypted = COALESCE(?, google_service_account_encrypted),
      google_admin_email = ?,
      google_customer = ?,
      timezone = ?,
      past_days = ?,
      future_days = ?,
      concurrency = ?,
      sync_interval_minutes = ?,
      sync_new_users_by_default = ?,
      enabled = ?,
      setup_completed = ?,
      updated_at = ?
      WHERE id = 1`)
    .bind(
      baseUrl || null,
      tokenEncrypted,
      serviceAccountEncrypted,
      adminEmail || null,
      customer || "my_customer",
      timezone,
      clampInteger(input.pastDays, current.pastDays, 0, 365),
      clampInteger(input.futureDays, current.futureDays, 1, 730),
      clampInteger(input.concurrency, current.concurrency, 1, 10),
      clampInteger(input.syncIntervalMinutes, current.syncIntervalMinutes, 15, 1440),
      input.syncNewUsersByDefault === undefined ? Number(current.syncNewUsersByDefault) : Number(input.syncNewUsersByDefault),
      input.enabled === undefined ? Number(current.enabled) : Number(input.enabled),
      input.setupCompleted === undefined ? Number(current.setupCompleted) : Number(input.setupCompleted),
      now,
    )
    .run();

  await addAudit(actor, "configuration.updated", "Connection or sync settings were updated");
  return getConfig(false);
}

export async function addAudit(actor: string, action: string, detail?: string): Promise<void> {
  await ensureSchema();
  await db()
    .prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, ?, ?)")
    .bind(new Date().toISOString(), actor, action, detail ?? null)
    .run();
}

export async function createRun(trigger: string): Promise<RunSummary> {
  await ensureSchema();
  const run: RunSummary = {
    id: crypto.randomUUID(),
    trigger,
    status: "running",
    startedAt: new Date().toISOString(),
    completedAt: null,
    usersDiscovered: 0,
    usersMatched: 0,
    usersSynced: 0,
    eventsCreated: 0,
    eventsUpdated: 0,
    eventsDeleted: 0,
    eventsUnchanged: 0,
    errors: 0,
    message: null,
  };
  const database = db();
  database.transaction(() => {
    const active = database
      .prepare("SELECT id FROM sync_runs WHERE status = 'running' LIMIT 1")
      .first<{ id: string }>();
    if (active) throw new HttpError(409, "A sync is already running", active.id);
    database
      .prepare("INSERT INTO sync_runs (id, trigger, status, started_at, heartbeat_at) VALUES (?, ?, 'running', ?, ?)")
      .bind(run.id, run.trigger, run.startedAt, run.startedAt)
      .run();
  });
  return run;
}

export async function touchRunHeartbeat(runId: string): Promise<void> {
  await ensureSchema();
  db().prepare("UPDATE sync_runs SET heartbeat_at = ? WHERE id = ? AND status = 'running'")
    .bind(new Date().toISOString(), runId).run();
}

export async function recoverStaleRuns(maxAgeMinutes = 5): Promise<number> {
  await ensureSchema();
  const result = db()
    .prepare(`UPDATE sync_runs SET status = 'failed', completed_at = ?, errors = errors + 1,
      message = 'Run was interrupted by a server restart or exceeded the maximum runtime.'
      WHERE status = 'running' AND COALESCE(heartbeat_at, started_at) <= ?`)
    .bind(
      new Date().toISOString(),
      new Date(Date.now() - maxAgeMinutes * 60_000).toISOString(),
    )
    .run();
  return Number(result.changes);
}

export async function finishRun(run: RunSummary): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`UPDATE sync_runs SET status = ?, completed_at = ?, heartbeat_at = ?, users_discovered = ?, users_matched = ?,
      users_synced = ?, events_created = ?, events_updated = ?, events_deleted = ?, events_unchanged = ?,
      errors = ?, message = ? WHERE id = ?`)
    .bind(
      run.status,
      run.completedAt,
      run.completedAt,
      run.usersDiscovered,
      run.usersMatched,
      run.usersSynced,
      run.eventsCreated,
      run.eventsUpdated,
      run.eventsDeleted,
      run.eventsUnchanged,
      run.errors,
      run.message,
      run.id,
    )
    .run();
}

export async function listRuns(limit = 30): Promise<RunSummary[]> {
  await ensureSchema();
  const result = await db()
    .prepare(`SELECT id, trigger, status, started_at AS startedAt, completed_at AS completedAt,
      users_discovered AS usersDiscovered, users_matched AS usersMatched, users_synced AS usersSynced,
      events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
      events_unchanged AS eventsUnchanged, errors, message FROM sync_runs ORDER BY started_at DESC LIMIT ?`)
    .bind(Math.max(1, Math.min(limit, 100)))
    .all<RunSummary>();
  return result.results;
}

type UserMappingWrite = Omit<UserMapping, "syncEnabled" | "directoryActive">;

export async function upsertUserMapping(mapping: UserMappingWrite): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`INSERT INTO user_mappings
      (google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name, role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 1, ?)
      ON CONFLICT(google_user_id) DO UPDATE SET google_email=excluded.google_email, schoolbox_user_id=excluded.schoolbox_user_id,
      schoolbox_email=excluded.schoolbox_email, display_name=excluded.display_name, role=excluded.role, status=excluded.status,
      last_sync_at=excluded.last_sync_at, last_error=excluded.last_error, event_count=excluded.event_count,
      directory_active=1, updated_at=excluded.updated_at`)
    .bind(
      mapping.googleUserId,
      mapping.googleEmail,
      mapping.schoolboxUserId,
      mapping.schoolboxEmail,
      mapping.displayName,
      mapping.role,
      mapping.status,
      mapping.lastSyncAt,
      mapping.lastError,
      mapping.eventCount,
      mapping.updatedAt,
    )
    .run();
}

/**
 * Records a complete, successful directory discovery and returns each user's
 * persisted selection. The organisation default is used only for new rows;
 * an administrator's existing selection is deliberately absent from the
 * conflict update.
 */
export async function discoverUserMappings(
  discoveries: UserMappingWrite[],
  defaultEnabled: boolean,
): Promise<Map<string, boolean>> {
  await ensureSchema();
  const binding = db();
  const upsert = binding.prepare(`INSERT INTO user_mappings
    (google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name, role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
    ON CONFLICT(google_user_id) DO UPDATE SET
      google_email=excluded.google_email,
      schoolbox_user_id=excluded.schoolbox_user_id,
      schoolbox_email=excluded.schoolbox_email,
      display_name=excluded.display_name,
      role=excluded.role,
      status=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN 'unmatched'
        WHEN user_mappings.status = 'unmatched' THEN 'pending'
        ELSE user_mappings.status
      END,
      last_sync_at=user_mappings.last_sync_at,
      last_error=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN excluded.last_error
        WHEN user_mappings.status = 'unmatched' THEN NULL
        ELSE user_mappings.last_error
      END,
      event_count=user_mappings.event_count,
      directory_active=1,
      updated_at=excluded.updated_at`);
  const selection = binding.prepare("SELECT sync_enabled FROM user_mappings WHERE google_user_id = ?");

  return binding.transaction(() => {
    const result = new Map<string, boolean>();
    const discoveryTime = discoveries[0]?.updatedAt ?? new Date().toISOString();
    // Email addresses are mutable. Retiring the prior snapshot first makes
    // swaps and reassignment order-independent while the stable Google ID keeps
    // each user's selection and event mappings attached to the correct row.
    binding.prepare("UPDATE user_mappings SET directory_active = 0, updated_at = ? WHERE directory_active = 1")
      .bind(discoveryTime).run();
    for (const mapping of discoveries) {
      upsert.bind(
        mapping.googleUserId,
        mapping.googleEmail,
        mapping.schoolboxUserId,
        mapping.schoolboxEmail,
        mapping.displayName,
        mapping.role,
        mapping.status,
        mapping.lastSyncAt,
        mapping.lastError,
        mapping.eventCount,
        Number(defaultEnabled),
        mapping.updatedAt,
      ).run();
      const row = selection.bind(mapping.googleUserId).first<{ sync_enabled: number }>();
      result.set(mapping.googleUserId, Boolean(row?.sync_enabled));
    }
    return result;
  });
}

export async function setUsersSyncEnabled(ids: string[], enabled: boolean, actor: string): Promise<number> {
  await ensureSchema();
  const uniqueIds = [...new Set(ids.map((id) => id.trim()).filter(Boolean))];
  if (uniqueIds.length === 0) throw new HttpError(400, "Choose at least one user");
  if (uniqueIds.length > 25_000) throw new HttpError(400, "Update no more than 25,000 users at a time");

  const binding = db();
  const exists = binding.prepare("SELECT google_user_id FROM user_mappings WHERE google_user_id = ? AND directory_active = 1");
  const update = binding.prepare("UPDATE user_mappings SET sync_enabled = ?, updated_at = ? WHERE google_user_id = ? AND directory_active = 1");
  const audit = binding.prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, 'users.sync_selection_updated', ?)");
  const now = new Date().toISOString();
  const updated = binding.transaction(() => {
    for (const id of uniqueIds) {
      if (!exists.bind(id).first()) throw new HttpError(404, "One or more users are no longer available");
    }
    let changes = 0;
    for (const id of uniqueIds) {
      changes += Number(update.bind(Number(enabled), now, id).run().changes);
    }
    audit.bind(now, actor, `${uniqueIds.length} user selection(s) set to ${enabled ? "enabled" : "paused"}`).run();
    return changes;
  });
  return updated;
}

export async function listUserMappings(limit?: number, includeInactive = false): Promise<UserMapping[]> {
  await ensureSchema();
  const statement = db()
    .prepare(`SELECT u.google_user_id AS googleUserId, u.google_email AS googleEmail, u.schoolbox_user_id AS schoolboxUserId,
      u.schoolbox_email AS schoolboxEmail, u.display_name AS displayName, u.role, u.status, u.last_sync_at AS lastSyncAt,
      u.last_error AS lastError,
      (SELECT COUNT(*) FROM event_mappings e WHERE e.google_user_id = u.google_user_id) AS eventCount,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM user_mappings u${includeInactive ? "" : " WHERE u.directory_active = 1"}
      ORDER BY u.google_email${limit === undefined ? "" : " LIMIT ?"}`);
  const result = limit === undefined
    ? statement.all<UserMapping>()
    : statement.bind(Math.max(1, Math.min(limit, 5000))).all<UserMapping>();
  for (const mapping of result.results) {
    mapping.syncEnabled = Boolean(mapping.syncEnabled);
    mapping.directoryActive = Boolean(mapping.directoryActive);
  }
  return result.results;
}

export async function getUserMapping(googleUserId: string): Promise<UserMapping | null> {
  await ensureSchema();
  const mapping = db()
    .prepare(`SELECT u.google_user_id AS googleUserId, u.google_email AS googleEmail, u.schoolbox_user_id AS schoolboxUserId,
      u.schoolbox_email AS schoolboxEmail, u.display_name AS displayName, u.role, u.status, u.last_sync_at AS lastSyncAt,
      u.last_error AS lastError,
      (SELECT COUNT(*) FROM event_mappings e WHERE e.google_user_id = u.google_user_id) AS eventCount,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM user_mappings u WHERE u.google_user_id = ? AND u.directory_active = 1`)
    .bind(googleUserId)
    .first<UserMapping>();
  if (mapping) {
    mapping.syncEnabled = Boolean(mapping.syncEnabled);
    mapping.directoryActive = Boolean(mapping.directoryActive);
  }
  return mapping;
}

export async function getEventMappings(googleUserId: string): Promise<EventMapping[]> {
  await ensureSchema();
  const result = await db()
    .prepare(`SELECT google_user_id AS googleUserId, source_key AS sourceKey, google_event_id AS googleEventId,
      source_hash AS sourceHash, source_start AS sourceStart, source_end AS sourceEnd, last_seen_run_id AS lastSeenRunId,
      created_at AS createdAt, updated_at AS updatedAt FROM event_mappings WHERE google_user_id = ?`)
    .bind(googleUserId)
    .all<EventMapping>();
  return result.results;
}

export async function upsertEventMapping(mapping: EventMapping): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`INSERT INTO event_mappings
      (google_user_id, source_key, google_event_id, source_hash, source_start, source_end, last_seen_run_id, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(google_user_id, source_key) DO UPDATE SET google_event_id=excluded.google_event_id,
      source_hash=excluded.source_hash, source_start=excluded.source_start, source_end=excluded.source_end,
      last_seen_run_id=excluded.last_seen_run_id, updated_at=excluded.updated_at`)
    .bind(
      mapping.googleUserId,
      mapping.sourceKey,
      mapping.googleEventId,
      mapping.sourceHash,
      mapping.sourceStart,
      mapping.sourceEnd,
      mapping.lastSeenRunId,
      mapping.createdAt,
      mapping.updatedAt,
    )
    .run();
}

export async function touchEventMapping(googleUserId: string, sourceKey: string, runId: string): Promise<void> {
  await ensureSchema();
  await db()
    .prepare("UPDATE event_mappings SET last_seen_run_id = ?, updated_at = ? WHERE google_user_id = ? AND source_key = ?")
    .bind(runId, new Date().toISOString(), googleUserId, sourceKey)
    .run();
}

export async function deleteEventMapping(googleUserId: string, sourceKey: string): Promise<void> {
  await ensureSchema();
  await db().prepare("DELETE FROM event_mappings WHERE google_user_id = ? AND source_key = ?").bind(googleUserId, sourceKey).run();
}

export async function recordManagedEventCleanup(options: {
  googleUserId: string;
  remaining: number;
  deleted: number;
  alreadyMissing: number;
  error: string | null;
  actor: string;
}): Promise<void> {
  await ensureSchema();
  if (!Number.isInteger(options.remaining) || options.remaining < 0) {
    throw new HttpError(400, "The remaining event count is invalid");
  }
  const now = new Date().toISOString();
  const error = options.error?.slice(0, 2_000) ?? null;
  const binding = db();
  binding.transaction(() => {
    const result = binding.prepare(`UPDATE user_mappings SET sync_enabled = 0, event_count = ?,
      status = CASE WHEN ? IS NOT NULL THEN 'error' WHEN schoolbox_user_id IS NULL THEN 'unmatched' ELSE 'pending' END,
      last_error = ?, updated_at = ? WHERE google_user_id = ? AND directory_active = 1`)
      .bind(options.remaining, error, error, now, options.googleUserId)
      .run();
    if (Number(result.changes) !== 1) throw new HttpError(404, "This user is no longer available");
    binding.prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, 'users.managed_events_cleanup', ?)")
      .bind(
        now,
        options.actor,
        `${options.deleted} managed event(s) deleted, ${options.alreadyMissing} already absent, ${options.remaining} remaining`,
      )
      .run();
  });
}

export async function statusSnapshot(): Promise<{
  configured: boolean;
  config: AppConfig;
  lastRun: RunSummary | null;
  counts: { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number };
}> {
  const config = await getConfig(false);
  const [runs, userCounts, events] = await Promise.all([
    listRuns(1),
    db().prepare(`SELECT SUM(CASE WHEN directory_active = 1 THEN 1 ELSE 0 END) AS users,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 THEN 1 ELSE 0 END) AS enabled,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 0 THEN 1 ELSE 0 END) AS disabled,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'synced' THEN 1 ELSE 0 END) AS healthy,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'error' THEN 1 ELSE 0 END) AS errors,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'unmatched' THEN 1 ELSE 0 END) AS unmatched
      FROM user_mappings`).first<{ users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number }>(),
    db().prepare(`SELECT COUNT(*) AS count FROM event_mappings e
      JOIN user_mappings u ON u.google_user_id = e.google_user_id WHERE u.directory_active = 1`).first<{ count: number }>(),
  ]);
  return {
    configured: Boolean(config.schoolboxBaseUrl && config.hasSchoolboxToken && config.hasGoogleServiceAccount && config.googleAdminEmail),
    config,
    lastRun: runs[0] ?? null,
    counts: {
      users: userCounts?.users ?? 0,
      enabled: userCounts?.enabled ?? 0,
      disabled: userCounts?.disabled ?? 0,
      healthy: userCounts?.healthy ?? 0,
      errors: userCounts?.errors ?? 0,
      unmatched: userCounts?.unmatched ?? 0,
      events: events?.count ?? 0,
    },
  };
}
