import { db } from "./db";
import {
  DEFAULT_SYNC_POLICY,
  EVENT_CATEGORIES,
  eventTypeKey,
  normalizeEventTypeLabel,
  normalizeSyncPolicy,
  normalizeUserEventExclusions,
  type EventCategory,
  type SyncPolicy,
  type SyncPolicyInput,
  type UserEventExclusions,
  type UserEventExclusionsInput,
} from "./policy";
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
  discoveryTimeoutSeconds: number;
  userSyncTimeoutSeconds: number;
  runTimeoutMinutes: number;
  syncIntervalMinutes: number;
  syncNewUsersByDefault: boolean;
  syncPolicy: SyncPolicy;
  enabled: boolean;
  setupCompleted: boolean;
  hasSchoolboxToken: boolean;
  hasGoogleServiceAccount: boolean;
  serviceAccountEmail?: string;
  serviceAccountClientId?: string;
  updatedAt?: string;
};

export type ConfigInput = Partial<Omit<AppConfig, "hasSchoolboxToken" | "hasGoogleServiceAccount" | "serviceAccountEmail" | "serviceAccountClientId" | "syncPolicy">> & {
  syncPolicy?: SyncPolicyInput;
};

export type DiscoveredEventType = {
  key: string;
  label: string;
  category: EventCategory;
  lastSeenAt: string;
};

export type RunSummary = {
  id: string;
  trigger: string;
  status: string;
  phase: string | null;
  phaseDetail: string | null;
  progressAt: string | null;
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
  calendarCount: number;
  hasCustomExclusions: boolean;
  syncEnabled: boolean;
  directoryActive: boolean;
  updatedAt: string;
};

export type EventMapping = {
  googleUserId: string;
  sourceKey: string;
  googleEventId: string;
  calendarId: string;
  sourceHash: string;
  sourceStart: string;
  sourceEnd: string;
  lastSeenRunId: string;
  createdAt: string;
  updatedAt: string;
  title: string | null;
  description: string | null;
  location: string | null;
  author: string | null;
  eventType: string | null;
  category: string | null;
  allDay: boolean;
  sourceUrl: string | null;
  destinationId: string | null;
};

export type RunUserDiagnostic = {
  runId: string;
  googleUserId: string;
  googleEmail: string;
  displayName: string | null;
  schoolboxUserId: number | null;
  schoolboxEmail: string | null;
  status: string;
  stage: string;
  startedAt: string;
  completedAt: string | null;
  eventsFound: number;
  eventsIncluded: number;
  eventsCreated: number;
  eventsUpdated: number;
  eventsDeleted: number;
  eventsUnchanged: number;
  managedEventsAfter: number;
  errorMessage: string | null;
};

export type RunEventDiagnostic = {
  runId: string;
  googleUserId: string;
  sourceKey: string;
  title: string | null;
  description: string | null;
  location: string | null;
  author: string | null;
  eventType: string | null;
  category: string | null;
  sourceStart: string | null;
  sourceEnd: string | null;
  allDay: boolean;
  sourceUrl: string | null;
  googleEventId: string | null;
  calendarId: string | null;
  destinationId: string | null;
  action: string;
  detail: string | null;
  errorMessage: string | null;
  recordedAt: string;
};

export type UserCalendarTarget = {
  googleUserId: string;
  destinationId: string;
  googleCalendarId: string;
  summary: string;
  description: string;
  timeZone: string;
  createdAt: string;
  updatedAt: string;
};

export type UserCalendarTargetWithEmail = UserCalendarTarget & {
  googleEmail: string;
};

export type CalendarDestinationUsage = {
  destinationId: string;
  summary: string;
  calendarCount: number;
  eventCount: number;
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
    discovery_timeout_seconds INTEGER NOT NULL DEFAULT 120,
    user_sync_timeout_seconds INTEGER NOT NULL DEFAULT 180,
    run_timeout_minutes INTEGER NOT NULL DEFAULT 30,
    sync_interval_minutes INTEGER NOT NULL DEFAULT 360,
    sync_new_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_users_by_default IN (0, 1)),
    sync_policy_json TEXT NOT NULL DEFAULT '{}',
    enabled INTEGER NOT NULL DEFAULT 0,
    setup_completed INTEGER NOT NULL DEFAULT 0,
    updated_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS sync_runs (
    id TEXT PRIMARY KEY,
    trigger TEXT NOT NULL,
    status TEXT NOT NULL,
    phase TEXT,
    phase_detail TEXT,
    progress_at TEXT,
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
    calendar_id TEXT NOT NULL DEFAULT 'primary',
    source_hash TEXT NOT NULL,
    source_start TEXT NOT NULL,
    source_end TEXT NOT NULL,
    last_seen_run_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    title TEXT,
    description TEXT,
    location TEXT,
    author TEXT,
    event_type TEXT,
    category TEXT,
    all_day INTEGER NOT NULL DEFAULT 0 CHECK (all_day IN (0, 1)),
    source_url TEXT,
    destination_id TEXT,
    PRIMARY KEY (google_user_id, source_key)
  )`,
  `CREATE TABLE IF NOT EXISTS sync_run_users (
    run_id TEXT NOT NULL,
    google_user_id TEXT NOT NULL,
    google_email TEXT NOT NULL,
    display_name TEXT,
    schoolbox_user_id INTEGER,
    schoolbox_email TEXT,
    status TEXT NOT NULL,
    stage TEXT NOT NULL,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    events_found INTEGER NOT NULL DEFAULT 0,
    events_included INTEGER NOT NULL DEFAULT 0,
    events_created INTEGER NOT NULL DEFAULT 0,
    events_updated INTEGER NOT NULL DEFAULT 0,
    events_deleted INTEGER NOT NULL DEFAULT 0,
    events_unchanged INTEGER NOT NULL DEFAULT 0,
    managed_events_after INTEGER NOT NULL DEFAULT 0,
    error_message TEXT,
    PRIMARY KEY (run_id, google_user_id)
  )`,
  `CREATE TABLE IF NOT EXISTS sync_run_events (
    run_id TEXT NOT NULL,
    google_user_id TEXT NOT NULL,
    source_key TEXT NOT NULL,
    title TEXT,
    description TEXT,
    location TEXT,
    author TEXT,
    event_type TEXT,
    category TEXT,
    source_start TEXT,
    source_end TEXT,
    all_day INTEGER NOT NULL DEFAULT 0 CHECK (all_day IN (0, 1)),
    source_url TEXT,
    google_event_id TEXT,
    calendar_id TEXT,
    destination_id TEXT,
    action TEXT NOT NULL,
    detail TEXT,
    error_message TEXT,
    recorded_at TEXT NOT NULL,
    PRIMARY KEY (run_id, google_user_id, source_key)
  )`,
  `CREATE TABLE IF NOT EXISTS user_calendar_targets (
    google_user_id TEXT NOT NULL,
    destination_id TEXT NOT NULL,
    google_calendar_id TEXT NOT NULL,
    summary TEXT NOT NULL,
    description TEXT NOT NULL DEFAULT '',
    time_zone TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (google_user_id, destination_id)
  )`,
  `CREATE TABLE IF NOT EXISTS event_type_catalog (
    type_key TEXT PRIMARY KEY,
    label TEXT NOT NULL,
    category TEXT NOT NULL,
    last_seen_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS user_event_exclusions (
    google_user_id TEXT PRIMARY KEY,
    excluded_categories_json TEXT NOT NULL DEFAULT '[]',
    excluded_event_types_json TEXT NOT NULL DEFAULT '[]',
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL
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
  "CREATE INDEX IF NOT EXISTS sync_run_users_status_idx ON sync_run_users (run_id, status, google_email)",
  "CREATE INDEX IF NOT EXISTS sync_run_events_user_idx ON sync_run_events (run_id, google_user_id, action, source_start)",
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
  if (!runColumns.some((column) => column.name === "phase")) {
    binding.prepare("ALTER TABLE sync_runs ADD COLUMN phase TEXT").run();
  }
  if (!runColumns.some((column) => column.name === "phase_detail")) {
    binding.prepare("ALTER TABLE sync_runs ADD COLUMN phase_detail TEXT").run();
  }
  if (!runColumns.some((column) => column.name === "progress_at")) {
    binding.prepare("ALTER TABLE sync_runs ADD COLUMN progress_at TEXT").run();
  }
  const configColumns = binding.prepare("PRAGMA table_info(app_config)").all<{ name: string }>().results;
  if (!configColumns.some((column) => column.name === "sync_new_users_by_default")) {
    // Legacy installations implicitly synced every newly discovered account. Keep
    // that behaviour on upgrade; brand-new databases use the safer CREATE default.
    binding.prepare("ALTER TABLE app_config ADD COLUMN sync_new_users_by_default INTEGER NOT NULL DEFAULT 1 CHECK (sync_new_users_by_default IN (0, 1))").run();
  }
  if (!configColumns.some((column) => column.name === "sync_policy_json")) {
    // Existing installations synchronized every category and copied all content.
    // The policy defaults preserve that behaviour until an administrator changes it.
    binding.prepare("ALTER TABLE app_config ADD COLUMN sync_policy_json TEXT NOT NULL DEFAULT '{}'").run();
  }
  if (!configColumns.some((column) => column.name === "discovery_timeout_seconds")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN discovery_timeout_seconds INTEGER NOT NULL DEFAULT 120").run();
  }
  if (!configColumns.some((column) => column.name === "user_sync_timeout_seconds")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN user_sync_timeout_seconds INTEGER NOT NULL DEFAULT 180").run();
  }
  if (!configColumns.some((column) => column.name === "run_timeout_minutes")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN run_timeout_minutes INTEGER NOT NULL DEFAULT 30").run();
  }
  const eventMappingColumns = binding.prepare("PRAGMA table_info(event_mappings)").all<{ name: string }>().results;
  if (!eventMappingColumns.some((column) => column.name === "calendar_id")) {
    // Every event created by older Relay versions was placed on the primary calendar.
    binding.prepare("ALTER TABLE event_mappings ADD COLUMN calendar_id TEXT NOT NULL DEFAULT 'primary'").run();
  }
  const eventDiagnosticColumns: Array<{ name: string; definition: string }> = [
    { name: "title", definition: "TEXT" },
    { name: "description", definition: "TEXT" },
    { name: "location", definition: "TEXT" },
    { name: "author", definition: "TEXT" },
    { name: "event_type", definition: "TEXT" },
    { name: "category", definition: "TEXT" },
    { name: "all_day", definition: "INTEGER NOT NULL DEFAULT 0 CHECK (all_day IN (0, 1))" },
    { name: "source_url", definition: "TEXT" },
    { name: "destination_id", definition: "TEXT" },
  ];
  for (const column of eventDiagnosticColumns) {
    if (!eventMappingColumns.some((existing) => existing.name === column.name)) {
      binding.prepare(`ALTER TABLE event_mappings ADD COLUMN ${column.name} ${column.definition}`).run();
    }
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
  // An unmatched Google account has no Schoolbox identity to sync. Keep this
  // invariant in storage so legacy data and direct API calls cannot make the
  // account appear enabled or inflate the paused-user total.
  binding.prepare("UPDATE user_mappings SET sync_enabled = 0 WHERE schoolbox_user_id IS NULL AND sync_enabled <> 0").run();
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
  discovery_timeout_seconds: number;
  user_sync_timeout_seconds: number;
  run_timeout_minutes: number;
  sync_interval_minutes: number;
  sync_new_users_by_default: number;
  sync_policy_json: string;
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
    discoveryTimeoutSeconds: row.discovery_timeout_seconds,
    userSyncTimeoutSeconds: row.user_sync_timeout_seconds,
    runTimeoutMinutes: row.run_timeout_minutes,
    syncIntervalMinutes: row.sync_interval_minutes,
    syncNewUsersByDefault: Boolean(row.sync_new_users_by_default),
    syncPolicy: (() => {
      try { return normalizeSyncPolicy(JSON.parse(row.sync_policy_json || "{}"), DEFAULT_SYNC_POLICY); }
      catch { return normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY); }
    })(),
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
  try { new Intl.DateTimeFormat("en-AU", { timeZone: timezone }).format(new Date()); }
  catch { throw new HttpError(400, "Enter a valid IANA calendar time zone"); }
  if (input.syncPolicy?.secondaryCalendars !== undefined) {
    if (!Array.isArray(input.syncPolicy.secondaryCalendars)) throw new HttpError(400, "Secondary calendar destinations must be a list");
    const names = input.syncPolicy.secondaryCalendars.map((calendar) => calendar?.name?.trim().toLocaleLowerCase("en-AU") ?? "");
    if (names.some((name) => !name)) throw new HttpError(400, "Give every secondary calendar destination a name");
    if (new Set(names).size !== names.length) throw new HttpError(400, "Each secondary calendar destination needs a unique name");
  }
  const syncPolicy = normalizeSyncPolicy(input.syncPolicy ?? {}, current.syncPolicy);
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
      discovery_timeout_seconds = ?,
      user_sync_timeout_seconds = ?,
      run_timeout_minutes = ?,
      sync_interval_minutes = ?,
      sync_new_users_by_default = ?,
      sync_policy_json = ?,
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
      clampInteger(input.discoveryTimeoutSeconds, current.discoveryTimeoutSeconds, 30, 900),
      clampInteger(input.userSyncTimeoutSeconds, current.userSyncTimeoutSeconds, 30, 1800),
      clampInteger(input.runTimeoutMinutes, current.runTimeoutMinutes, 5, 240),
      clampInteger(input.syncIntervalMinutes, current.syncIntervalMinutes, 15, 1440),
      input.syncNewUsersByDefault === undefined ? Number(current.syncNewUsersByDefault) : Number(input.syncNewUsersByDefault),
      JSON.stringify(syncPolicy),
      input.enabled === undefined ? Number(current.enabled) : Number(input.enabled),
      input.setupCompleted === undefined ? Number(current.setupCompleted) : Number(input.setupCompleted),
      now,
    )
    .run();

  await addAudit(actor, "configuration.updated", "Connection or sync settings were updated");
  return getConfig(false);
}

export async function recordDiscoveredEventTypes(events: Array<{ type: string | null; category: EventCategory }>): Promise<void> {
  await ensureSchema();
  const now = new Date().toISOString();
  const discovered = new Map<string, { label: string; category: EventCategory }>();
  for (const event of events) {
    const label = normalizeEventTypeLabel(event.type);
    const key = eventTypeKey(label);
    if (!key || !EVENT_CATEGORIES.includes(event.category)) continue;
    discovered.set(key, { label, category: event.category });
  }
  if (discovered.size === 0) return;
  db().transaction(() => {
    const statement = db().prepare(`INSERT INTO event_type_catalog (type_key, label, category, last_seen_at)
      VALUES (?, ?, ?, ?)
      ON CONFLICT(type_key) DO UPDATE SET
        label = excluded.label,
        category = CASE WHEN event_type_catalog.category = excluded.category THEN excluded.category ELSE 'other' END,
        last_seen_at = excluded.last_seen_at`);
    for (const [key, event] of discovered) statement.bind(key, event.label, event.category, now).run();
  });
}

export async function listDiscoveredEventTypes(): Promise<DiscoveredEventType[]> {
  await ensureSchema();
  return db().prepare(`SELECT type_key AS key, label, category, last_seen_at AS lastSeenAt
    FROM event_type_catalog ORDER BY label COLLATE NOCASE`).all<DiscoveredEventType>().results;
}

export async function getUserEventExclusions(googleUserId: string): Promise<UserEventExclusions> {
  await ensureSchema();
  const row = db().prepare(`SELECT excluded_categories_json AS categoriesJson,
    excluded_event_types_json AS eventTypesJson, updated_at AS updatedAt, updated_by AS updatedBy
    FROM user_event_exclusions WHERE google_user_id = ?`)
    .bind(googleUserId)
    .first<{ categoriesJson: string; eventTypesJson: string; updatedAt: string; updatedBy: string }>();
  if (!row) return normalizeUserEventExclusions(null);
  try {
    return normalizeUserEventExclusions({
      categories: JSON.parse(row.categoriesJson),
      eventTypes: JSON.parse(row.eventTypesJson),
    }, row);
  } catch {
    return normalizeUserEventExclusions(null, row);
  }
}

export async function saveUserEventExclusions(
  googleUserId: string,
  input: UserEventExclusionsInput,
  actor: string,
): Promise<UserEventExclusions> {
  await ensureSchema();
  const id = googleUserId.trim();
  if (!id || id.length > 200) throw new HttpError(400, "Choose a valid user");
  const binding = db();
  const user = binding.prepare(`SELECT schoolbox_user_id AS schoolboxUserId FROM user_mappings
    WHERE google_user_id = ? AND directory_active = 1`)
    .bind(id)
    .first<{ schoolboxUserId: number | null }>();
  if (!user) throw new HttpError(404, "User not found");
  if (user.schoolboxUserId === null) {
    throw new HttpError(409, "Event exclusions can only be configured after a Schoolbox identity is matched");
  }

  const now = new Date().toISOString();
  const normalized = normalizeUserEventExclusions(input, { updatedAt: now, updatedBy: actor });
  binding.transaction(() => {
    if (normalized.categories.length === 0 && normalized.eventTypes.length === 0) {
      binding.prepare("DELETE FROM user_event_exclusions WHERE google_user_id = ?").bind(id).run();
    } else {
      binding.prepare(`INSERT INTO user_event_exclusions
        (google_user_id, excluded_categories_json, excluded_event_types_json, updated_at, updated_by)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(google_user_id) DO UPDATE SET
          excluded_categories_json = excluded.excluded_categories_json,
          excluded_event_types_json = excluded.excluded_event_types_json,
          updated_at = excluded.updated_at,
          updated_by = excluded.updated_by`)
        .bind(id, JSON.stringify(normalized.categories), JSON.stringify(normalized.eventTypes), now, actor)
        .run();
    }
    binding.prepare(`INSERT INTO audit_log (occurred_at, actor, action, detail)
      VALUES (?, ?, 'user.event_exclusions_updated', ?)`)
      .bind(
        now,
        actor,
        `${id}: ${normalized.categories.length} category exclusion(s), ${normalized.eventTypes.length} exact type exclusion(s)`,
      )
      .run();
  });
  return normalized.categories.length === 0 && normalized.eventTypes.length === 0
    ? normalizeUserEventExclusions(null)
    : normalized;
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
    phase: "starting",
    phaseDetail: "Preparing synchronization clients.",
    progressAt: new Date().toISOString(),
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
      .prepare(`INSERT INTO sync_runs
        (id, trigger, status, phase, phase_detail, progress_at, started_at, heartbeat_at)
        VALUES (?, ?, 'running', ?, ?, ?, ?, ?)`)
      .bind(run.id, run.trigger, run.phase, run.phaseDetail, run.progressAt, run.startedAt, run.startedAt)
      .run();
  });
  return run;
}

export async function touchRunHeartbeat(runId: string): Promise<void> {
  await ensureSchema();
  db().prepare("UPDATE sync_runs SET heartbeat_at = ? WHERE id = ? AND status = 'running'")
    .bind(new Date().toISOString(), runId).run();
}

export async function checkpointRun(run: RunSummary, phase: string, detail: string): Promise<void> {
  await ensureSchema();
  const now = new Date().toISOString();
  run.phase = phase;
  run.phaseDetail = detail;
  run.progressAt = now;
  db().prepare(`UPDATE sync_runs SET phase = ?, phase_detail = ?, progress_at = ?, heartbeat_at = ?,
    users_discovered = ?, users_matched = ?, users_synced = ?, events_created = ?, events_updated = ?,
    events_deleted = ?, events_unchanged = ?, errors = ? WHERE id = ? AND status = 'running'`)
    .bind(
      phase,
      detail.slice(0, 2_000),
      now,
      now,
      run.usersDiscovered,
      run.usersMatched,
      run.usersSynced,
      run.eventsCreated,
      run.eventsUpdated,
      run.eventsDeleted,
      run.eventsUnchanged,
      run.errors,
      run.id,
    )
    .run();
}

export async function recoverStaleRuns(maxAgeMinutes = 5): Promise<number> {
  await ensureSchema();
  const result = db()
    .prepare(`UPDATE sync_runs SET status = 'failed', phase = 'failed',
      phase_detail = 'Run heartbeat stopped before completion.', completed_at = ?, errors = errors + 1,
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
  run.phase = run.status === "completed" || run.status === "completed_with_errors" ? "completed" : "failed";
  run.phaseDetail = run.message;
  run.progressAt = run.completedAt;
  await db()
    .prepare(`UPDATE sync_runs SET status = ?, phase = ?, phase_detail = ?, progress_at = ?,
      completed_at = ?, heartbeat_at = ?, users_discovered = ?, users_matched = ?,
      users_synced = ?, events_created = ?, events_updated = ?, events_deleted = ?, events_unchanged = ?,
      errors = ?, message = ? WHERE id = ?`)
    .bind(
      run.status,
      run.phase,
      run.phaseDetail,
      run.progressAt,
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
  // Keep summaries indefinitely but bound sensitive, high-volume drill-down
  // snapshots to the newest 100 runs (roughly 25 days at a six-hour cadence).
  const expiredRuns = `SELECT id FROM sync_runs ORDER BY started_at DESC LIMIT -1 OFFSET 100`;
  db().prepare(`DELETE FROM sync_run_events WHERE run_id IN (${expiredRuns})`).run();
  db().prepare(`DELETE FROM sync_run_users WHERE run_id IN (${expiredRuns})`).run();
}

export async function listRuns(limit = 30): Promise<RunSummary[]> {
  await ensureSchema();
  const result = await db()
    .prepare(`SELECT id, trigger, status, started_at AS startedAt, completed_at AS completedAt,
      phase, phase_detail AS phaseDetail, progress_at AS progressAt,
      users_discovered AS usersDiscovered, users_matched AS usersMatched, users_synced AS usersSynced,
      events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
      events_unchanged AS eventsUnchanged, errors, message FROM sync_runs ORDER BY started_at DESC LIMIT ?`)
    .bind(Math.max(1, Math.min(limit, 100)))
    .all<RunSummary>();
  return result.results;
}

export async function getRun(runId: string): Promise<RunSummary | null> {
  await ensureSchema();
  return db().prepare(`SELECT id, trigger, status, started_at AS startedAt, completed_at AS completedAt,
    phase, phase_detail AS phaseDetail, progress_at AS progressAt,
    users_discovered AS usersDiscovered, users_matched AS usersMatched, users_synced AS usersSynced,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, errors, message FROM sync_runs WHERE id = ?`)
    .bind(runId).first<RunSummary>();
}

export async function startRunUserDiagnostic(input: Omit<RunUserDiagnostic,
  "status" | "stage" | "startedAt" | "completedAt" | "eventsFound" | "eventsIncluded" |
  "eventsCreated" | "eventsUpdated" | "eventsDeleted" | "eventsUnchanged" |
  "managedEventsAfter" | "errorMessage">): Promise<void> {
  await ensureSchema();
  const startedAt = new Date().toISOString();
  db().prepare(`INSERT INTO sync_run_users
    (run_id, google_user_id, google_email, display_name, schoolbox_user_id, schoolbox_email,
     status, stage, started_at)
    VALUES (?, ?, ?, ?, ?, ?, 'running', 'starting', ?)
    ON CONFLICT(run_id, google_user_id) DO UPDATE SET
      google_email=excluded.google_email, display_name=excluded.display_name,
      schoolbox_user_id=excluded.schoolbox_user_id, schoolbox_email=excluded.schoolbox_email,
      status='running', stage='starting', started_at=excluded.started_at, completed_at=NULL,
      events_found=0, events_included=0, events_created=0, events_updated=0,
      events_deleted=0, events_unchanged=0, managed_events_after=0, error_message=NULL`)
    .bind(
      input.runId,
      input.googleUserId,
      input.googleEmail,
      input.displayName,
      input.schoolboxUserId,
      input.schoolboxEmail,
      startedAt,
    ).run();
}

export async function finishRunUserDiagnostic(input: Pick<RunUserDiagnostic,
  "runId" | "googleUserId" | "status" | "stage" | "eventsFound" | "eventsIncluded" |
  "eventsCreated" | "eventsUpdated" | "eventsDeleted" | "eventsUnchanged" |
  "managedEventsAfter" | "errorMessage">): Promise<void> {
  await ensureSchema();
  db().prepare(`UPDATE sync_run_users SET status = ?, stage = ?, completed_at = ?,
    events_found = ?, events_included = ?, events_created = ?, events_updated = ?,
    events_deleted = ?, events_unchanged = ?, managed_events_after = ?, error_message = ?
    WHERE run_id = ? AND google_user_id = ?`)
    .bind(
      input.status,
      input.stage,
      new Date().toISOString(),
      input.eventsFound,
      input.eventsIncluded,
      input.eventsCreated,
      input.eventsUpdated,
      input.eventsDeleted,
      input.eventsUnchanged,
      input.managedEventsAfter,
      input.errorMessage?.slice(0, 4_000) ?? null,
      input.runId,
      input.googleUserId,
    ).run();
}

export async function recordRunEventDiagnostic(input: Omit<RunEventDiagnostic, "recordedAt">): Promise<void> {
  await ensureSchema();
  db().prepare(`INSERT INTO sync_run_events
    (run_id, google_user_id, source_key, title, description, location, author, event_type,
     category, source_start, source_end, all_day, source_url, google_event_id, calendar_id,
     destination_id, action, detail, error_message, recorded_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(run_id, google_user_id, source_key) DO UPDATE SET
      title=excluded.title, description=excluded.description, location=excluded.location,
      author=excluded.author, event_type=excluded.event_type, category=excluded.category,
      source_start=excluded.source_start, source_end=excluded.source_end, all_day=excluded.all_day,
      source_url=excluded.source_url, google_event_id=excluded.google_event_id,
      calendar_id=excluded.calendar_id, destination_id=excluded.destination_id,
      action=excluded.action, detail=excluded.detail, error_message=excluded.error_message,
      recorded_at=excluded.recorded_at`)
    .bind(
      input.runId,
      input.googleUserId,
      input.sourceKey,
      input.title,
      input.description,
      input.location,
      input.author,
      input.eventType,
      input.category,
      input.sourceStart,
      input.sourceEnd,
      Number(input.allDay),
      input.sourceUrl,
      input.googleEventId,
      input.calendarId,
      input.destinationId,
      input.action,
      input.detail?.slice(0, 2_000) ?? null,
      input.errorMessage?.slice(0, 4_000) ?? null,
      new Date().toISOString(),
    ).run();
}

export async function listRunUserDiagnostics(runId: string): Promise<RunUserDiagnostic[]> {
  await ensureSchema();
  return db().prepare(`SELECT run_id AS runId, google_user_id AS googleUserId,
    google_email AS googleEmail, display_name AS displayName, schoolbox_user_id AS schoolboxUserId,
    schoolbox_email AS schoolboxEmail, status, stage, started_at AS startedAt,
    completed_at AS completedAt, events_found AS eventsFound, events_included AS eventsIncluded,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, managed_events_after AS managedEventsAfter,
    error_message AS errorMessage FROM sync_run_users WHERE run_id = ?
    ORDER BY CASE status WHEN 'failed' THEN 0 WHEN 'running' THEN 1 ELSE 2 END,
      google_email COLLATE NOCASE`)
    .bind(runId).all<RunUserDiagnostic>().results;
}

export async function listRunEventDiagnostics(
  runId: string,
  googleUserId: string,
  options: { limit?: number; offset?: number } = {},
): Promise<{ events: RunEventDiagnostic[]; total: number }> {
  await ensureSchema();
  const limit = Math.max(1, Math.min(options.limit ?? 100, 250));
  const offset = Math.max(0, options.offset ?? 0);
  const binding = db();
  const total = binding.prepare(`SELECT COUNT(*) AS count FROM sync_run_events
    WHERE run_id = ? AND google_user_id = ?`).bind(runId, googleUserId).first<{ count: number }>()?.count ?? 0;
  const events = binding.prepare(`SELECT run_id AS runId, google_user_id AS googleUserId,
    source_key AS sourceKey, title, description, location, author, event_type AS eventType,
    category, source_start AS sourceStart, source_end AS sourceEnd, all_day AS allDay,
    source_url AS sourceUrl, google_event_id AS googleEventId, calendar_id AS calendarId,
    destination_id AS destinationId, action, detail, error_message AS errorMessage,
    recorded_at AS recordedAt FROM sync_run_events
    WHERE run_id = ? AND google_user_id = ?
    ORDER BY CASE action WHEN 'failed' THEN 0 WHEN 'created' THEN 1 WHEN 'updated' THEN 2
      WHEN 'deleted' THEN 3 ELSE 4 END, source_start, title COLLATE NOCASE LIMIT ? OFFSET ?`)
    .bind(runId, googleUserId, limit, offset).all<RunEventDiagnostic>().results;
  for (const event of events) event.allDay = Boolean(event.allDay);
  return { events, total };
}

export async function listUserRunDiagnostics(googleUserId: string, limit = 20): Promise<RunUserDiagnostic[]> {
  await ensureSchema();
  return db().prepare(`SELECT run_id AS runId, google_user_id AS googleUserId,
    google_email AS googleEmail, display_name AS displayName, schoolbox_user_id AS schoolboxUserId,
    schoolbox_email AS schoolboxEmail, status, stage, started_at AS startedAt,
    completed_at AS completedAt, events_found AS eventsFound, events_included AS eventsIncluded,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, managed_events_after AS managedEventsAfter,
    error_message AS errorMessage FROM sync_run_users WHERE google_user_id = ?
    ORDER BY started_at DESC LIMIT ?`)
    .bind(googleUserId, Math.max(1, Math.min(limit, 100))).all<RunUserDiagnostic>().results;
}

type UserMappingWrite = Omit<UserMapping, "syncEnabled" | "directoryActive" | "calendarCount" | "hasCustomExclusions">;

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
      sync_enabled=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN 0
        ELSE user_mappings.sync_enabled
      END,
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
        Number(defaultEnabled && mapping.schoolboxUserId !== null),
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
  const exists = binding.prepare("SELECT google_user_id, schoolbox_user_id FROM user_mappings WHERE google_user_id = ? AND directory_active = 1");
  const update = binding.prepare("UPDATE user_mappings SET sync_enabled = ?, updated_at = ? WHERE google_user_id = ? AND directory_active = 1");
  const audit = binding.prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, 'users.sync_selection_updated', ?)");
  const now = new Date().toISOString();
  const updated = binding.transaction(() => {
    for (const id of uniqueIds) {
      const user = exists.bind(id).first<{ google_user_id: string; schoolbox_user_id: string | null }>();
      if (!user) throw new HttpError(404, "One or more users are no longer available");
      if (enabled && user.schoolbox_user_id === null) {
        throw new HttpError(409, "Unmatched users cannot be enabled until a Schoolbox identity is discovered");
      }
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
      (SELECT COUNT(*) FROM user_calendar_targets c WHERE c.google_user_id = u.google_user_id) AS calendarCount,
      EXISTS(SELECT 1 FROM user_event_exclusions x WHERE x.google_user_id = u.google_user_id) AS hasCustomExclusions,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM user_mappings u${includeInactive ? "" : " WHERE u.directory_active = 1"}
      ORDER BY u.google_email${limit === undefined ? "" : " LIMIT ?"}`);
  const result = limit === undefined
    ? statement.all<UserMapping>()
    : statement.bind(Math.max(1, Math.min(limit, 5000))).all<UserMapping>();
  for (const mapping of result.results) {
    mapping.hasCustomExclusions = Boolean(mapping.hasCustomExclusions);
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
      (SELECT COUNT(*) FROM user_calendar_targets c WHERE c.google_user_id = u.google_user_id) AS calendarCount,
      EXISTS(SELECT 1 FROM user_event_exclusions x WHERE x.google_user_id = u.google_user_id) AS hasCustomExclusions,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM user_mappings u WHERE u.google_user_id = ? AND u.directory_active = 1`)
    .bind(googleUserId)
    .first<UserMapping>();
  if (mapping) {
    mapping.hasCustomExclusions = Boolean(mapping.hasCustomExclusions);
    mapping.syncEnabled = Boolean(mapping.syncEnabled);
    mapping.directoryActive = Boolean(mapping.directoryActive);
  }
  return mapping;
}

export async function getEventMappings(googleUserId: string): Promise<EventMapping[]> {
  await ensureSchema();
  const result = await db()
    .prepare(`SELECT google_user_id AS googleUserId, source_key AS sourceKey, google_event_id AS googleEventId,
      calendar_id AS calendarId, source_hash AS sourceHash, source_start AS sourceStart,
      source_end AS sourceEnd, last_seen_run_id AS lastSeenRunId, created_at AS createdAt,
      updated_at AS updatedAt, title, description, location, author, event_type AS eventType,
      category, all_day AS allDay, source_url AS sourceUrl, destination_id AS destinationId
      FROM event_mappings WHERE google_user_id = ? ORDER BY source_start, title COLLATE NOCASE`)
    .bind(googleUserId)
    .all<EventMapping>();
  for (const event of result.results) event.allDay = Boolean(event.allDay);
  return result.results;
}

type EventMappingDiagnosticFields = "title" | "description" | "location" | "author" |
  "eventType" | "category" | "allDay" | "sourceUrl" | "destinationId";

export async function upsertEventMapping(
  mapping: Omit<EventMapping, "calendarId" | EventMappingDiagnosticFields> &
    { calendarId?: string } & Partial<Pick<EventMapping, EventMappingDiagnosticFields>>,
): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`INSERT INTO event_mappings
      (google_user_id, source_key, google_event_id, calendar_id, source_hash, source_start, source_end,
       last_seen_run_id, created_at, updated_at, title, description, location, author, event_type,
       category, all_day, source_url, destination_id)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(google_user_id, source_key) DO UPDATE SET google_event_id=excluded.google_event_id,
      calendar_id=excluded.calendar_id, source_hash=excluded.source_hash, source_start=excluded.source_start, source_end=excluded.source_end,
      last_seen_run_id=excluded.last_seen_run_id, updated_at=excluded.updated_at,
      title=excluded.title, description=excluded.description, location=excluded.location,
      author=excluded.author, event_type=excluded.event_type, category=excluded.category,
      all_day=excluded.all_day, source_url=excluded.source_url, destination_id=excluded.destination_id`)
    .bind(
      mapping.googleUserId,
      mapping.sourceKey,
      mapping.googleEventId,
      mapping.calendarId?.trim() || "primary",
      mapping.sourceHash,
      mapping.sourceStart,
      mapping.sourceEnd,
      mapping.lastSeenRunId,
      mapping.createdAt,
      mapping.updatedAt,
      mapping.title ?? null,
      mapping.description ?? null,
      mapping.location ?? null,
      mapping.author ?? null,
      mapping.eventType ?? null,
      mapping.category ?? null,
      Number(mapping.allDay ?? false),
      mapping.sourceUrl ?? null,
      mapping.destinationId ?? null,
    )
    .run();
}

export async function getUserCalendarTarget(
  googleUserId: string,
  destinationId: string,
): Promise<UserCalendarTarget | null> {
  await ensureSchema();
  return db().prepare(`SELECT google_user_id AS googleUserId, destination_id AS destinationId,
    google_calendar_id AS googleCalendarId, summary, description, time_zone AS timeZone,
    created_at AS createdAt, updated_at AS updatedAt
    FROM user_calendar_targets WHERE google_user_id = ? AND destination_id = ?`)
    .bind(googleUserId, destinationId)
    .first<UserCalendarTarget>();
}

export async function listUserCalendarTargets(googleUserId: string): Promise<UserCalendarTarget[]> {
  await ensureSchema();
  return db().prepare(`SELECT google_user_id AS googleUserId, destination_id AS destinationId,
    google_calendar_id AS googleCalendarId, summary, description, time_zone AS timeZone,
    created_at AS createdAt, updated_at AS updatedAt
    FROM user_calendar_targets WHERE google_user_id = ? ORDER BY destination_id`)
    .bind(googleUserId)
    .all<UserCalendarTarget>().results;
}

export async function listCalendarDestinationUsage(): Promise<CalendarDestinationUsage[]> {
  await ensureSchema();
  return db().prepare(`SELECT t.destination_id AS destinationId, MAX(t.summary) AS summary,
    COUNT(DISTINCT t.google_user_id) AS calendarCount, COUNT(e.source_key) AS eventCount
    FROM user_calendar_targets t
    LEFT JOIN event_mappings e ON e.google_user_id = t.google_user_id AND e.calendar_id = t.google_calendar_id
    GROUP BY t.destination_id ORDER BY MAX(t.summary) COLLATE NOCASE`)
    .all<CalendarDestinationUsage>().results;
}

export async function listCalendarTargetsForDestination(destinationId: string): Promise<UserCalendarTargetWithEmail[]> {
  await ensureSchema();
  return db().prepare(`SELECT t.google_user_id AS googleUserId, u.google_email AS googleEmail,
    t.destination_id AS destinationId, t.google_calendar_id AS googleCalendarId,
    t.summary, t.description, t.time_zone AS timeZone, t.created_at AS createdAt, t.updated_at AS updatedAt
    FROM user_calendar_targets t
    JOIN user_mappings u ON u.google_user_id = t.google_user_id
    WHERE t.destination_id = ? ORDER BY u.google_email COLLATE NOCASE`)
    .bind(destinationId)
    .all<UserCalendarTargetWithEmail>().results;
}

export async function deleteCalendarTargetRecords(
  googleUserId: string,
  destinationId: string,
  googleCalendarId: string,
): Promise<number> {
  await ensureSchema();
  const binding = db();
  return binding.transaction(() => {
    const target = binding.prepare(`SELECT 1 AS found FROM user_calendar_targets
      WHERE google_user_id = ? AND destination_id = ? AND google_calendar_id = ?`)
      .bind(googleUserId, destinationId, googleCalendarId)
      .first<{ found: number }>();
    if (!target) return 0;
    const removedEvents = Number(binding.prepare(`DELETE FROM event_mappings
      WHERE google_user_id = ? AND calendar_id = ?`)
      .bind(googleUserId, googleCalendarId).run().changes);
    binding.prepare(`DELETE FROM user_calendar_targets
      WHERE google_user_id = ? AND destination_id = ? AND google_calendar_id = ?`)
      .bind(googleUserId, destinationId, googleCalendarId).run();
    return removedEvents;
  });
}

export async function upsertUserCalendarTarget(target: UserCalendarTarget): Promise<void> {
  await ensureSchema();
  await db().prepare(`INSERT INTO user_calendar_targets
    (google_user_id, destination_id, google_calendar_id, summary, description, time_zone, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(google_user_id, destination_id) DO UPDATE SET
      google_calendar_id=excluded.google_calendar_id, summary=excluded.summary,
      description=excluded.description, time_zone=excluded.time_zone, updated_at=excluded.updated_at`)
    .bind(
      target.googleUserId,
      target.destinationId,
      target.googleCalendarId,
      target.summary,
      target.description,
      target.timeZone,
      target.createdAt,
      target.updatedAt,
    )
    .run();
}

export async function touchEventMapping(
  googleUserId: string,
  sourceKey: string,
  runId: string,
  diagnostic: Partial<Pick<EventMapping, EventMappingDiagnosticFields>> = {},
): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`UPDATE event_mappings SET last_seen_run_id = ?, updated_at = ?,
      title = COALESCE(?, title), description = COALESCE(?, description),
      location = COALESCE(?, location), author = COALESCE(?, author),
      event_type = COALESCE(?, event_type), category = COALESCE(?, category),
      all_day = ?, source_url = COALESCE(?, source_url),
      destination_id = COALESCE(?, destination_id)
      WHERE google_user_id = ? AND source_key = ?`)
    .bind(
      runId,
      new Date().toISOString(),
      diagnostic.title ?? null,
      diagnostic.description ?? null,
      diagnostic.location ?? null,
      diagnostic.author ?? null,
      diagnostic.eventType ?? null,
      diagnostic.category ?? null,
      Number(diagnostic.allDay ?? false),
      diagnostic.sourceUrl ?? null,
      diagnostic.destinationId ?? null,
      googleUserId,
      sourceKey,
    )
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
  calendarsDeleted?: number;
  calendarsAlreadyMissing?: number;
  calendarsRemaining?: number;
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
        `${options.deleted} managed event(s) deleted, ${options.alreadyMissing} already absent, ${options.remaining} remaining; ${options.calendarsDeleted ?? 0} managed calendar(s) deleted, ${options.calendarsAlreadyMissing ?? 0} already absent, ${options.calendarsRemaining ?? 0} remaining`,
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
      SUM(CASE WHEN directory_active = 1 AND schoolbox_user_id IS NOT NULL AND sync_enabled = 1 THEN 1 ELSE 0 END) AS enabled,
      SUM(CASE WHEN directory_active = 1 AND schoolbox_user_id IS NOT NULL AND sync_enabled = 0 THEN 1 ELSE 0 END) AS disabled,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'synced' THEN 1 ELSE 0 END) AS healthy,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'error' THEN 1 ELSE 0 END) AS errors,
      SUM(CASE WHEN directory_active = 1 AND status = 'unmatched' THEN 1 ELSE 0 END) AS unmatched
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
