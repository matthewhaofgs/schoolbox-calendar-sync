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
  googleEnabled: boolean;
  googleSetupCompleted: boolean;
  syncNewGoogleUsersByDefault: boolean;
  syncNewUsersByDefault: boolean;
  microsoftTenantId: string;
  microsoftClientId: string;
  microsoftClientSecret?: string;
  microsoftTestUserEmail: string;
  microsoftEnabled: boolean;
  microsoftSetupCompleted: boolean;
  syncNewMicrosoftUsersByDefault: boolean;
  microsoftSyncPolicy: SyncPolicy;
  microsoftConsentGrantedAt: string;
  syncPolicy: SyncPolicy;
  enabled: boolean;
  setupCompleted: boolean;
  schoolboxSetupCompleted: boolean;
  schoolboxConfigured: boolean;
  googleConfigured: boolean;
  microsoftConfigured: boolean;
  hasSchoolboxToken: boolean;
  hasGoogleServiceAccount: boolean;
  hasMicrosoftClientSecret: boolean;
  serviceAccountEmail?: string;
  serviceAccountClientId?: string;
  updatedAt?: string;
};

export type ConfigInput = Partial<Omit<AppConfig, "hasSchoolboxToken" | "hasGoogleServiceAccount" | "hasMicrosoftClientSecret" | "serviceAccountEmail" | "serviceAccountClientId" | "schoolboxConfigured" | "googleConfigured" | "microsoftConfigured" | "setupCompleted" | "syncPolicy" | "microsoftSyncPolicy">> & {
  syncPolicy?: SyncPolicyInput;
  microsoftSyncPolicy?: SyncPolicyInput;
};

export const TARGET_PROVIDERS = ["google", "microsoft"] as const;
export type TargetProvider = (typeof TARGET_PROVIDERS)[number];

export function normalizeTargetProvider(value: unknown): TargetProvider {
  if (value === "google" || value === "microsoft") return value;
  throw new HttpError(400, "Choose Google Workspace or Microsoft 365");
}

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
  target: TargetProvider;
  targetUserId: string;
  targetEmail: string;
  /** Legacy aliases retained for Google-compatible API consumers. */
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
  target: TargetProvider;
  targetUserId: string;
  targetEventId: string;
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
  target: TargetProvider;
  targetUserId: string;
  targetEmail: string;
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
  target: TargetProvider;
  targetUserId: string;
  targetEventId: string | null;
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
  target: TargetProvider;
  targetUserId: string;
  targetCalendarId: string;
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
  targetEmail: string;
  googleEmail: string;
};

export type CalendarDestinationUsage = {
  target: TargetProvider;
  destinationId: string;
  summary: string;
  calendarCount: number;
  eventCount: number;
};

export type RunTargetSummary = {
  runId: string;
  target: TargetProvider;
  status: string;
  phase: string;
  phaseDetail: string | null;
  startedAt: string;
  completedAt: string | null;
  usersDiscovered: number;
  usersMatched: number;
  usersSelected: number;
  usersSynced: number;
  eventsCreated: number;
  eventsUpdated: number;
  eventsDeleted: number;
  eventsUnchanged: number;
  errors: number;
  message: string | null;
};

const schemaStatements = [
  `CREATE TABLE IF NOT EXISTS app_config (
    id INTEGER PRIMARY KEY DEFAULT 1,
    schoolbox_base_url TEXT,
    schoolbox_token_encrypted TEXT,
    schoolbox_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (schoolbox_setup_completed IN (0, 1)),
    schoolbox_credential_version INTEGER NOT NULL DEFAULT 0,
    schoolbox_verified_version INTEGER NOT NULL DEFAULT -1,
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
    google_sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK (google_sync_enabled IN (0, 1)),
    google_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (google_setup_completed IN (0, 1)),
    google_credential_version INTEGER NOT NULL DEFAULT 0,
    google_verified_version INTEGER NOT NULL DEFAULT -1,
    sync_new_google_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_google_users_by_default IN (0, 1)),
    microsoft_tenant_id TEXT,
    microsoft_client_id TEXT,
    microsoft_client_secret_encrypted TEXT,
    microsoft_credential_version INTEGER NOT NULL DEFAULT 0,
    microsoft_test_user_email TEXT,
    microsoft_sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK (microsoft_sync_enabled IN (0, 1)),
    microsoft_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (microsoft_setup_completed IN (0, 1)),
    microsoft_verified_version INTEGER NOT NULL DEFAULT -1,
    sync_new_microsoft_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_microsoft_users_by_default IN (0, 1)),
    microsoft_sync_policy_json TEXT NOT NULL DEFAULT '{}',
    microsoft_consent_granted_at TEXT,
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
  `CREATE TABLE IF NOT EXISTS target_user_mappings (
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    target_user_id TEXT NOT NULL,
    target_email TEXT NOT NULL,
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
    updated_at TEXT NOT NULL,
    PRIMARY KEY (target, target_user_id)
  )`,
  `CREATE TABLE IF NOT EXISTS target_event_mappings (
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    target_user_id TEXT NOT NULL,
    source_key TEXT NOT NULL,
    target_event_id TEXT NOT NULL,
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
    PRIMARY KEY (target, target_user_id, source_key)
  )`,
  `CREATE TABLE IF NOT EXISTS target_calendar_targets (
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    target_user_id TEXT NOT NULL,
    destination_id TEXT NOT NULL,
    target_calendar_id TEXT NOT NULL,
    summary TEXT NOT NULL,
    description TEXT NOT NULL DEFAULT '',
    time_zone TEXT NOT NULL,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    PRIMARY KEY (target, target_user_id, destination_id)
  )`,
  `CREATE TABLE IF NOT EXISTS schoolbox_user_exclusions (
    schoolbox_user_id INTEGER PRIMARY KEY,
    excluded_categories_json TEXT NOT NULL DEFAULT '[]',
    excluded_event_types_json TEXT NOT NULL DEFAULT '[]',
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS sync_run_targets (
    run_id TEXT NOT NULL,
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    status TEXT NOT NULL,
    phase TEXT NOT NULL,
    phase_detail TEXT,
    started_at TEXT NOT NULL,
    completed_at TEXT,
    users_discovered INTEGER NOT NULL DEFAULT 0,
    users_matched INTEGER NOT NULL DEFAULT 0,
    users_selected INTEGER NOT NULL DEFAULT 0,
    users_synced INTEGER NOT NULL DEFAULT 0,
    events_created INTEGER NOT NULL DEFAULT 0,
    events_updated INTEGER NOT NULL DEFAULT 0,
    events_deleted INTEGER NOT NULL DEFAULT 0,
    events_unchanged INTEGER NOT NULL DEFAULT 0,
    errors INTEGER NOT NULL DEFAULT 0,
    message TEXT,
    PRIMARY KEY (run_id, target)
  )`,
  `CREATE TABLE IF NOT EXISTS sync_run_target_users (
    run_id TEXT NOT NULL,
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    target_user_id TEXT NOT NULL,
    target_email TEXT NOT NULL,
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
    PRIMARY KEY (run_id, target, target_user_id)
  )`,
  `CREATE TABLE IF NOT EXISTS sync_run_target_events (
    run_id TEXT NOT NULL,
    target TEXT NOT NULL CHECK (target IN ('google', 'microsoft')),
    target_user_id TEXT NOT NULL,
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
    target_event_id TEXT,
    calendar_id TEXT,
    destination_id TEXT,
    action TEXT NOT NULL,
    detail TEXT,
    error_message TEXT,
    recorded_at TEXT NOT NULL,
    PRIMARY KEY (run_id, target, target_user_id, source_key)
  )`,
  `CREATE TABLE IF NOT EXISTS microsoft_consent_states (
    state_hash TEXT PRIMARY KEY,
    tenant_id TEXT NOT NULL,
    client_id TEXT NOT NULL,
    credential_version INTEGER NOT NULL,
    actor TEXT NOT NULL,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL
  )`,
  `CREATE TABLE IF NOT EXISTS schema_meta (
    id INTEGER PRIMARY KEY CHECK (id = 1),
    version INTEGER NOT NULL,
    updated_at TEXT NOT NULL
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
  "CREATE UNIQUE INDEX IF NOT EXISTS target_user_mappings_active_email_idx ON target_user_mappings (target, target_email COLLATE NOCASE) WHERE directory_active = 1",
  "CREATE INDEX IF NOT EXISTS target_event_mappings_seen_idx ON target_event_mappings (target, target_user_id, last_seen_run_id)",
  "CREATE INDEX IF NOT EXISTS sync_run_target_users_status_idx ON sync_run_target_users (run_id, target, status, target_email)",
  "CREATE INDEX IF NOT EXISTS sync_run_target_events_user_idx ON sync_run_target_events (run_id, target, target_user_id, action, source_start)",
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
  if (!configColumns.some((column) => column.name === "schoolbox_setup_completed")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN schoolbox_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (schoolbox_setup_completed IN (0, 1))").run();
    binding.prepare(`UPDATE app_config SET schoolbox_setup_completed = 1
      WHERE setup_completed = 1
        AND schoolbox_base_url IS NOT NULL AND trim(schoolbox_base_url) <> ''
        AND schoolbox_token_encrypted IS NOT NULL`).run();
  }
  if (!configColumns.some((column) => column.name === "schoolbox_credential_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN schoolbox_credential_version INTEGER NOT NULL DEFAULT 0").run();
  }
  if (!configColumns.some((column) => column.name === "schoolbox_verified_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN schoolbox_verified_version INTEGER NOT NULL DEFAULT -1").run();
    binding.prepare("UPDATE app_config SET schoolbox_verified_version = schoolbox_credential_version WHERE schoolbox_setup_completed = 1").run();
  }
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
  if (!configColumns.some((column) => column.name === "google_sync_enabled")) {
    // Existing installations are Google deployments. Preserve that target on upgrade.
    binding.prepare("ALTER TABLE app_config ADD COLUMN google_sync_enabled INTEGER NOT NULL DEFAULT 1 CHECK (google_sync_enabled IN (0, 1))").run();
  }
  if (!configColumns.some((column) => column.name === "google_setup_completed")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN google_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (google_setup_completed IN (0, 1))").run();
    // The original setup flow only supported Google. Preserve completion for
    // upgraded installations whose legacy setup was complete and whose stored
    // Google connection is still structurally ready.
    binding.prepare(`UPDATE app_config SET google_setup_completed = 1
      WHERE setup_completed = 1
        AND google_service_account_encrypted IS NOT NULL
        AND google_admin_email IS NOT NULL AND trim(google_admin_email) <> ''`).run();
  }
  if (!configColumns.some((column) => column.name === "google_credential_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN google_credential_version INTEGER NOT NULL DEFAULT 0").run();
  }
  if (!configColumns.some((column) => column.name === "google_verified_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN google_verified_version INTEGER NOT NULL DEFAULT -1").run();
    binding.prepare("UPDATE app_config SET google_verified_version = google_credential_version WHERE google_setup_completed = 1").run();
  }
  if (!configColumns.some((column) => column.name === "sync_new_google_users_by_default")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN sync_new_google_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_google_users_by_default IN (0, 1))").run();
    binding.prepare("UPDATE app_config SET sync_new_google_users_by_default = sync_new_users_by_default").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_tenant_id")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_tenant_id TEXT").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_client_id")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_client_id TEXT").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_client_secret_encrypted")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_client_secret_encrypted TEXT").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_credential_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_credential_version INTEGER NOT NULL DEFAULT 0").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_test_user_email")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_test_user_email TEXT").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_sync_enabled")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_sync_enabled INTEGER NOT NULL DEFAULT 0 CHECK (microsoft_sync_enabled IN (0, 1))").run();
  }
  if (!configColumns.some((column) => column.name === "sync_new_microsoft_users_by_default")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN sync_new_microsoft_users_by_default INTEGER NOT NULL DEFAULT 0 CHECK (sync_new_microsoft_users_by_default IN (0, 1))").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_sync_policy_json")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_sync_policy_json TEXT NOT NULL DEFAULT '{}'").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_consent_granted_at")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_consent_granted_at TEXT").run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_setup_completed")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_setup_completed INTEGER NOT NULL DEFAULT 0 CHECK (microsoft_setup_completed IN (0, 1))").run();
    binding.prepare(`UPDATE app_config SET microsoft_setup_completed = 1
      WHERE setup_completed = 1
        AND microsoft_tenant_id IS NOT NULL AND trim(microsoft_tenant_id) <> ''
        AND microsoft_client_id IS NOT NULL AND trim(microsoft_client_id) <> ''
        AND microsoft_client_secret_encrypted IS NOT NULL
        AND microsoft_consent_granted_at IS NOT NULL`).run();
  }
  if (!configColumns.some((column) => column.name === "microsoft_verified_version")) {
    binding.prepare("ALTER TABLE app_config ADD COLUMN microsoft_verified_version INTEGER NOT NULL DEFAULT -1").run();
    binding.prepare("UPDATE app_config SET microsoft_verified_version = microsoft_credential_version WHERE microsoft_setup_completed = 1").run();
  }
  const consentStateColumns = binding.prepare("PRAGMA table_info(microsoft_consent_states)").all<{ name: string }>().results;
  if (!consentStateColumns.some((column) => column.name === "client_id")) {
    binding.prepare("ALTER TABLE microsoft_consent_states ADD COLUMN client_id TEXT NOT NULL DEFAULT ''").run();
  }
  if (!consentStateColumns.some((column) => column.name === "credential_version")) {
    binding.prepare("ALTER TABLE microsoft_consent_states ADD COLUMN credential_version INTEGER NOT NULL DEFAULT 0").run();
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

  // Migrate the original Google-only persistence exactly once. Replaying this
  // copy after cleanup or retirement would resurrect deliberately deleted rows
  // from the now-read-only compatibility tables on the next process restart.
  const schemaVersion = binding.prepare("SELECT version FROM schema_meta WHERE id = 1")
    .first<{ version: number }>()?.version ?? 0;
  if (schemaVersion < 2) binding.transaction(() => {
    binding.prepare(`INSERT OR IGNORE INTO target_user_mappings
      (target, target_user_id, target_email, schoolbox_user_id, schoolbox_email, display_name,
       role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
      SELECT 'google', google_user_id, google_email, schoolbox_user_id, schoolbox_email, display_name,
       role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at
      FROM user_mappings`).run();
    binding.prepare(`INSERT OR IGNORE INTO target_event_mappings
      (target, target_user_id, source_key, target_event_id, calendar_id, source_hash, source_start,
       source_end, last_seen_run_id, created_at, updated_at, title, description, location, author,
       event_type, category, all_day, source_url, destination_id)
      SELECT 'google', google_user_id, source_key, google_event_id, calendar_id, source_hash,
       source_start, source_end, last_seen_run_id, created_at, updated_at, title, description,
       location, author, event_type, category, all_day, source_url, destination_id
      FROM event_mappings`).run();
    binding.prepare(`INSERT OR IGNORE INTO target_calendar_targets
      (target, target_user_id, destination_id, target_calendar_id, summary, description,
       time_zone, created_at, updated_at)
      SELECT 'google', google_user_id, destination_id, google_calendar_id, summary, description,
       time_zone, created_at, updated_at FROM user_calendar_targets`).run();
    binding.prepare(`INSERT OR IGNORE INTO schoolbox_user_exclusions
      (schoolbox_user_id, excluded_categories_json, excluded_event_types_json, updated_at, updated_by)
      SELECT u.schoolbox_user_id, x.excluded_categories_json, x.excluded_event_types_json,
       x.updated_at, x.updated_by FROM user_event_exclusions x
       JOIN user_mappings u ON u.google_user_id = x.google_user_id
       WHERE u.schoolbox_user_id IS NOT NULL ORDER BY x.updated_at DESC`).run();
    binding.prepare(`INSERT OR IGNORE INTO sync_run_target_users
      (run_id, target, target_user_id, target_email, display_name, schoolbox_user_id, schoolbox_email,
       status, stage, started_at, completed_at, events_found, events_included, events_created,
       events_updated, events_deleted, events_unchanged, managed_events_after, error_message)
      SELECT run_id, 'google', google_user_id, google_email, display_name, schoolbox_user_id,
       schoolbox_email, status, stage, started_at, completed_at, events_found, events_included,
       events_created, events_updated, events_deleted, events_unchanged, managed_events_after,
       error_message FROM sync_run_users`).run();
    binding.prepare(`INSERT OR IGNORE INTO sync_run_target_events
      (run_id, target, target_user_id, source_key, title, description, location, author, event_type,
       category, source_start, source_end, all_day, source_url, target_event_id, calendar_id,
       destination_id, action, detail, error_message, recorded_at)
      SELECT run_id, 'google', google_user_id, source_key, title, description, location, author,
       event_type, category, source_start, source_end, all_day, source_url, google_event_id,
       calendar_id, destination_id, action, detail, error_message, recorded_at FROM sync_run_events`).run();
    binding.prepare(`INSERT INTO schema_meta (id, version, updated_at) VALUES (1, 2, ?)
      ON CONFLICT(id) DO UPDATE SET version = MAX(schema_meta.version, excluded.version),
      updated_at = excluded.updated_at`).bind(new Date().toISOString()).run();
  });
  await binding
    .prepare("INSERT OR IGNORE INTO app_config (id, updated_at) VALUES (1, ?)")
    .bind(new Date().toISOString())
    .run();
  initialized = true;
}

type ConfigRow = {
  schoolbox_base_url: string | null;
  schoolbox_token_encrypted: string | null;
  schoolbox_setup_completed: number;
  schoolbox_credential_version: number;
  schoolbox_verified_version: number;
  google_service_account_encrypted: string | null;
  google_admin_email: string | null;
  google_customer: string;
  google_credential_version: number;
  google_verified_version: number;
  timezone: string;
  past_days: number;
  future_days: number;
  concurrency: number;
  discovery_timeout_seconds: number;
  user_sync_timeout_seconds: number;
  run_timeout_minutes: number;
  sync_interval_minutes: number;
  sync_new_users_by_default: number;
  google_sync_enabled: number;
  google_setup_completed: number;
  sync_new_google_users_by_default: number;
  microsoft_tenant_id: string | null;
  microsoft_client_id: string | null;
  microsoft_client_secret_encrypted: string | null;
  microsoft_credential_version: number;
  microsoft_test_user_email: string | null;
  microsoft_sync_enabled: number;
  microsoft_setup_completed: number;
  microsoft_verified_version: number;
  sync_new_microsoft_users_by_default: number;
  microsoft_sync_policy_json: string;
  microsoft_consent_granted_at: string | null;
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

  let microsoftClientSecret: string | undefined;
  if (row.microsoft_client_secret_encrypted) {
    try {
      microsoftClientSecret = await decryptSecret(row.microsoft_client_secret_encrypted);
    } catch {
      if (includeSecrets) throw new HttpError(500, "Stored Microsoft credential could not be decrypted");
    }
  }

  const schoolboxConfigured = Boolean(row.schoolbox_base_url && row.schoolbox_token_encrypted);
  const googleConfigured = Boolean(row.google_service_account_encrypted && row.google_admin_email?.trim());
  const microsoftConfigured = Boolean(
    row.microsoft_tenant_id?.trim() && row.microsoft_client_id?.trim() && row.microsoft_client_secret_encrypted,
  );
  const googleSetupCompleted = Boolean(row.google_setup_completed) && googleConfigured &&
    row.google_verified_version === row.google_credential_version;
  const microsoftSetupCompleted = Boolean(row.microsoft_setup_completed) &&
    microsoftConfigured && Boolean(row.microsoft_consent_granted_at) &&
    row.microsoft_verified_version === row.microsoft_credential_version;
  // `setupCompleted` remains a compatibility aggregate. Provider setup is
  // independent: Relay is operational when Schoolbox and at least one enabled,
  // verified calendar target are ready.
  const schoolboxSetupCompleted = Boolean(row.schoolbox_setup_completed) && schoolboxConfigured &&
    row.schoolbox_verified_version === row.schoolbox_credential_version;
  const setupCompleted = schoolboxSetupCompleted && (
    (Boolean(row.google_sync_enabled) && googleSetupCompleted) ||
    (Boolean(row.microsoft_sync_enabled) && microsoftSetupCompleted)
  );

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
    googleEnabled: Boolean(row.google_sync_enabled),
    googleSetupCompleted,
    syncNewGoogleUsersByDefault: Boolean(row.sync_new_google_users_by_default),
    // Compatibility alias for pre-Microsoft clients.
    syncNewUsersByDefault: Boolean(row.sync_new_google_users_by_default),
    microsoftTenantId: row.microsoft_tenant_id ?? "",
    microsoftClientId: row.microsoft_client_id ?? "",
    microsoftTestUserEmail: row.microsoft_test_user_email ?? "",
    microsoftEnabled: Boolean(row.microsoft_sync_enabled),
    microsoftSetupCompleted,
    syncNewMicrosoftUsersByDefault: Boolean(row.sync_new_microsoft_users_by_default),
    microsoftSyncPolicy: (() => {
      try { return normalizeSyncPolicy(JSON.parse(row.microsoft_sync_policy_json || "{}"), DEFAULT_SYNC_POLICY); }
      catch { return normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY); }
    })(),
    microsoftConsentGrantedAt: row.microsoft_consent_granted_at ?? "",
    syncPolicy: (() => {
      try { return normalizeSyncPolicy(JSON.parse(row.sync_policy_json || "{}"), DEFAULT_SYNC_POLICY); }
      catch { return normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY); }
    })(),
    enabled: Boolean(row.enabled),
    setupCompleted,
    schoolboxSetupCompleted,
    schoolboxConfigured,
    googleConfigured,
    microsoftConfigured,
    hasSchoolboxToken: Boolean(row.schoolbox_token_encrypted),
    hasGoogleServiceAccount: Boolean(row.google_service_account_encrypted),
    hasMicrosoftClientSecret: Boolean(row.microsoft_client_secret_encrypted),
    serviceAccountEmail,
    serviceAccountClientId,
    updatedAt: row.updated_at,
  };

  if (includeSecrets) {
    if (row.schoolbox_token_encrypted) result.schoolboxToken = await decryptSecret(row.schoolbox_token_encrypted);
    result.googleServiceAccountJson = googleServiceAccountJson;
    result.microsoftClientSecret = microsoftClientSecret;
  }
  return result;
}

export async function getStoredSchoolboxConnection(): Promise<{ baseUrl: string; token?: string; credentialVersion: number }> {
  await ensureSchema();
  const row = await db()
    .prepare("SELECT schoolbox_base_url, schoolbox_token_encrypted, schoolbox_credential_version FROM app_config WHERE id = 1")
    .first<{ schoolbox_base_url: string | null; schoolbox_token_encrypted: string | null; schoolbox_credential_version: number }>();
  return {
    baseUrl: row?.schoolbox_base_url ?? "",
    token: row?.schoolbox_token_encrypted ? await decryptSecret(row.schoolbox_token_encrypted) : undefined,
    credentialVersion: row?.schoolbox_credential_version ?? 0,
  };
}

export async function getStoredGoogleConnection(): Promise<{
  serviceAccountJson?: string;
  adminEmail: string;
  customer: string;
  credentialVersion: number;
}> {
  await ensureSchema();
  const row = await db()
    .prepare(`SELECT google_service_account_encrypted, google_admin_email, google_customer,
      google_credential_version FROM app_config WHERE id = 1`)
    .first<{ google_service_account_encrypted: string | null; google_admin_email: string | null; google_customer: string; google_credential_version: number }>();
  return {
    serviceAccountJson: row?.google_service_account_encrypted
      ? await decryptSecret(row.google_service_account_encrypted)
      : undefined,
    adminEmail: row?.google_admin_email ?? "",
    customer: row?.google_customer ?? "my_customer",
    credentialVersion: row?.google_credential_version ?? 0,
  };
}

export async function getStoredMicrosoftConnection(): Promise<{
  tenantId: string;
  clientId: string;
  clientSecret?: string;
  credentialVersion: number;
  testUserEmail: string;
}> {
  await ensureSchema();
  const row = await db().prepare(`SELECT microsoft_tenant_id AS tenantId,
    microsoft_client_id AS clientId, microsoft_client_secret_encrypted AS clientSecretEncrypted,
    microsoft_credential_version AS credentialVersion,
    microsoft_test_user_email AS testUserEmail FROM app_config WHERE id = 1`)
    .first<{ tenantId: string | null; clientId: string | null; clientSecretEncrypted: string | null; credentialVersion: number; testUserEmail: string | null }>();
  return {
    tenantId: row?.tenantId ?? "",
    clientId: row?.clientId ?? "",
    clientSecret: row?.clientSecretEncrypted ? await decryptSecret(row.clientSecretEncrypted) : undefined,
    credentialVersion: row?.credentialVersion ?? 0,
    testUserEmail: row?.testUserEmail ?? "",
  };
}

export async function recordConnectionVerified(
  target: "schoolbox" | TargetProvider,
  actor: string,
  credentialVersion: number,
): Promise<string> {
  await ensureSchema();
  const now = new Date().toISOString();
  const columns = target === "schoolbox"
    ? { verified: "schoolbox_verified_version", credential: "schoolbox_credential_version" }
    : target === "google"
      ? { verified: "google_verified_version", credential: "google_credential_version" }
      : { verified: "microsoft_verified_version", credential: "microsoft_credential_version" };
  const result = db().prepare(`UPDATE app_config SET ${columns.verified} = ?, updated_at = ?
    WHERE id = 1 AND ${columns.credential} = ?`)
    .bind(credentialVersion, now, credentialVersion).run();
  if (Number(result.changes) !== 1) {
    throw new HttpError(409, `${target === "schoolbox" ? "Schoolbox" : target === "google" ? "Google Workspace" : "Microsoft 365"} settings changed during verification; test the saved connection again`);
  }
  await addAudit(actor, `${target}.connection_verified`, `${target} stored connection passed its diagnostic`);
  return now;
}

async function stateHash(value: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(value));
  return Buffer.from(digest).toString("hex");
}

export type MicrosoftCredentialSnapshot = {
  tenantId: string;
  clientId: string;
  credentialVersion: number;
};

export async function createMicrosoftConsentState(
  connection: MicrosoftCredentialSnapshot,
  actor: string,
): Promise<string> {
  await ensureSchema();
  const state = Buffer.from(crypto.getRandomValues(new Uint8Array(32))).toString("base64url");
  const hash = await stateHash(state);
  const now = new Date();
  const expires = new Date(now.getTime() + 10 * 60_000);
  const binding = db();
  binding.transaction(() => {
    binding.prepare("DELETE FROM microsoft_consent_states WHERE expires_at <= ?").bind(now.toISOString()).run();
    binding.prepare(`INSERT INTO microsoft_consent_states
      (state_hash, tenant_id, client_id, credential_version, actor, created_at, expires_at)
      VALUES (?, ?, ?, ?, ?, ?, ?)`)
      .bind(hash, connection.tenantId, connection.clientId, connection.credentialVersion,
        actor, now.toISOString(), expires.toISOString()).run();
  });
  return state;
}

export async function consumeMicrosoftConsentState(
  state: string,
  returnedTenantId?: string,
): Promise<{ actor: string } & MicrosoftCredentialSnapshot> {
  await ensureSchema();
  if (state.length < 32 || state.length > 128) throw new HttpError(400, "Microsoft admin-consent state is invalid");
  const hash = await stateHash(state);
  const binding = db();
  const row = binding.prepare(`SELECT tenant_id AS tenantId, client_id AS clientId,
    credential_version AS credentialVersion, actor, expires_at AS expiresAt
    FROM microsoft_consent_states WHERE state_hash = ?`).bind(hash)
    .first<{ tenantId: string; clientId: string; credentialVersion: number; actor: string; expiresAt: string }>();
  binding.prepare("DELETE FROM microsoft_consent_states WHERE state_hash = ?").bind(hash).run();
  if (!row || row.expiresAt <= new Date().toISOString()) throw new HttpError(400, "Microsoft admin-consent state expired or was already used");
  if (returnedTenantId && row.tenantId.toLowerCase() !== returnedTenantId.trim().toLowerCase()) {
    throw new HttpError(400, "Microsoft returned consent for a different tenant");
  }
  return {
    actor: row.actor,
    tenantId: row.tenantId,
    clientId: row.clientId,
    credentialVersion: row.credentialVersion,
  };
}

export async function recordMicrosoftAdminConsent(
  actor: string,
  verifiedConnection: MicrosoftCredentialSnapshot,
): Promise<string> {
  await ensureSchema();
  const now = new Date().toISOString();
  const result = db().prepare(`UPDATE app_config SET microsoft_consent_granted_at = ?,
    microsoft_verified_version = microsoft_credential_version, updated_at = ?
    WHERE id = 1
      AND lower(microsoft_tenant_id) = lower(?)
      AND lower(microsoft_client_id) = lower(?)
      AND microsoft_credential_version = ?
      AND microsoft_client_secret_encrypted IS NOT NULL`)
    .bind(now, now, verifiedConnection.tenantId, verifiedConnection.clientId,
      verifiedConnection.credentialVersion).run();
  if (Number(result.changes) !== 1) {
    throw new HttpError(409, "Microsoft credentials changed during verification; test the saved connection again");
  }
  await addAudit(actor, "microsoft.admin_consent_recorded", "Microsoft 365 tenant admin consent returned successfully");
  return now;
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
  const microsoftTenantId = (input.microsoftTenantId ?? current.microsoftTenantId).trim().toLowerCase();
  const microsoftClientId = (input.microsoftClientId ?? current.microsoftClientId).trim().toLowerCase();
  const microsoftTestUserEmail = (input.microsoftTestUserEmail ?? current.microsoftTestUserEmail).trim().toLowerCase();
  const timezone = (input.timezone ?? current.timezone).trim() || "Australia/Sydney";
  try { new Intl.DateTimeFormat("en-AU", { timeZone: timezone }).format(new Date()); }
  catch { throw new HttpError(400, "Enter a valid IANA calendar time zone"); }
  for (const candidate of [input.syncPolicy, input.microsoftSyncPolicy]) {
    if (candidate?.secondaryCalendars === undefined) continue;
    if (!Array.isArray(candidate.secondaryCalendars)) throw new HttpError(400, "Secondary calendar destinations must be a list");
    const names = candidate.secondaryCalendars.map((calendar) => calendar?.name?.trim().toLocaleLowerCase("en-AU") ?? "");
    if (names.some((name) => !name)) throw new HttpError(400, "Give every secondary calendar destination a name");
    if (new Set(names).size !== names.length) throw new HttpError(400, "Each secondary calendar destination needs a unique name");
  }
  const syncPolicy = normalizeSyncPolicy(input.syncPolicy ?? {}, current.syncPolicy);
  const microsoftSyncPolicy = normalizeSyncPolicy(input.microsoftSyncPolicy ?? {}, current.microsoftSyncPolicy);
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
  const microsoftClientSecretEncrypted = input.microsoftClientSecret?.trim()
    ? await encryptSecret(input.microsoftClientSecret.trim())
    : null;
  const schoolboxConnectionChanged = baseUrl !== current.schoolboxBaseUrl || Boolean(input.schoolboxToken?.trim());
  const googleConnectionChanged = adminEmail !== current.googleAdminEmail ||
    customer !== current.googleCustomer || Boolean(input.googleServiceAccountJson?.trim());
  const microsoftIdentityCredentialChanged =
    microsoftTenantId !== current.microsoftTenantId ||
    microsoftClientId !== current.microsoftClientId ||
    Boolean(input.microsoftClientSecret?.trim());
  const microsoftConnectionChanged = microsoftIdentityCredentialChanged ||
    microsoftTestUserEmail !== current.microsoftTestUserEmail;
  const effectiveMicrosoftConsent = microsoftIdentityCredentialChanged ? null : current.microsoftConsentGrantedAt;
  const verification = db().prepare(`SELECT schoolbox_credential_version AS schoolboxCredentialVersion,
    schoolbox_verified_version AS schoolboxVerifiedVersion,
    google_credential_version AS googleCredentialVersion,
    google_verified_version AS googleVerifiedVersion,
    microsoft_credential_version AS microsoftCredentialVersion,
    microsoft_verified_version AS microsoftVerifiedVersion
    FROM app_config WHERE id = 1`).first<{
      schoolboxCredentialVersion: number; schoolboxVerifiedVersion: number;
      googleCredentialVersion: number; googleVerifiedVersion: number;
      microsoftCredentialVersion: number; microsoftVerifiedVersion: number;
    }>();
  if (!verification) throw new HttpError(500, "Application configuration row is missing");
  const schoolboxSetupCompleted = schoolboxConnectionChanged
    ? false
    : input.schoolboxSetupCompleted ?? current.schoolboxSetupCompleted;
  const googleSetupCompleted = googleConnectionChanged
    ? false
    : input.googleSetupCompleted ?? current.googleSetupCompleted;
  const microsoftSetupCompleted = microsoftConnectionChanged
    ? false
    : input.microsoftSetupCompleted ?? current.microsoftSetupCompleted;
  if (input.schoolboxSetupCompleted === true &&
    (schoolboxConnectionChanged || verification.schoolboxVerifiedVersion !== verification.schoolboxCredentialVersion)) {
    throw new HttpError(409, "Test the saved Schoolbox connection before completing its setup");
  }
  if (input.googleSetupCompleted === true &&
    (googleConnectionChanged || verification.googleVerifiedVersion !== verification.googleCredentialVersion)) {
    throw new HttpError(409, "Test the saved Google Workspace connection before completing its setup");
  }
  if (input.microsoftSetupCompleted === true &&
    (microsoftConnectionChanged || verification.microsoftVerifiedVersion !== verification.microsoftCredentialVersion)) {
    throw new HttpError(409, "Test the saved Microsoft 365 connection before completing its setup");
  }
  // Changing a provider connection invalidates and pauses only that target.
  // The scheduler and the other target retain their state.
  const requestedGoogleEnabled = input.googleEnabled ??
    (Boolean(input.googleServiceAccountJson?.trim()) || current.googleEnabled);
  const requestedMicrosoftEnabled = input.microsoftEnabled ?? current.microsoftEnabled;
  const googleEnabled = googleSetupCompleted && requestedGoogleEnabled;
  const microsoftEnabled = microsoftSetupCompleted && requestedMicrosoftEnabled;
  const syncNewGoogleUsersByDefault = input.syncNewGoogleUsersByDefault
    ?? input.syncNewUsersByDefault
    ?? current.syncNewGoogleUsersByDefault;
  const syncNewMicrosoftUsersByDefault = input.syncNewMicrosoftUsersByDefault
    ?? current.syncNewMicrosoftUsersByDefault;

  const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
  if (microsoftTenantId && !guid.test(microsoftTenantId)) throw new HttpError(400, "Enter a valid Microsoft Entra tenant ID");
  if (microsoftClientId && !guid.test(microsoftClientId)) throw new HttpError(400, "Enter a valid Microsoft application client ID");
  if (microsoftTestUserEmail && !/^\S+@\S+\.\S+$/.test(microsoftTestUserEmail)) {
    throw new HttpError(400, "Enter a valid Microsoft 365 test mailbox address");
  }
  const schedulerEnabled = input.enabled ?? current.enabled;
  const hasSchoolboxToken = Boolean(input.schoolboxToken?.trim() || current.hasSchoolboxToken);
  const hasGoogleCredential = Boolean(input.googleServiceAccountJson?.trim() || current.hasGoogleServiceAccount);
  const hasMicrosoftCredential = Boolean(input.microsoftClientSecret?.trim() || current.hasMicrosoftClientSecret);
  if (baseUrl) {
    let schoolboxUrl: URL;
    try { schoolboxUrl = new URL(baseUrl); } catch { throw new HttpError(400, "Enter a valid Schoolbox URL"); }
    if (schoolboxUrl.protocol !== "https:") throw new HttpError(400, "Schoolbox must use HTTPS");
  }
  if (schoolboxSetupCompleted && (!baseUrl || !hasSchoolboxToken)) {
    throw new HttpError(400, "Complete the Schoolbox connection before marking its setup complete");
  }
  if (googleSetupCompleted && (!adminEmail || !hasGoogleCredential)) {
    throw new HttpError(400, "Complete the Google Workspace connection before marking its setup complete");
  }
  if (microsoftSetupCompleted && (!microsoftTenantId || !microsoftClientId || !hasMicrosoftCredential || !effectiveMicrosoftConsent)) {
    throw new HttpError(400, "Complete Microsoft credentials, consent, and testing before marking its setup complete");
  }
  const setupCompleted = schoolboxSetupCompleted && (
    (googleEnabled && googleSetupCompleted) || (microsoftEnabled && microsoftSetupCompleted)
  );
  if (input.enabled === true && !setupCompleted) {
    throw new HttpError(400, "Complete Schoolbox and enable at least one completed calendar target before starting the schedule");
  }

  await ensureSchema();
  await db()
    .prepare(`UPDATE app_config SET
      schoolbox_base_url = ?,
      schoolbox_token_encrypted = COALESCE(?, schoolbox_token_encrypted),
      schoolbox_setup_completed = ?,
      schoolbox_credential_version = schoolbox_credential_version + ?,
      google_service_account_encrypted = COALESCE(?, google_service_account_encrypted),
      google_admin_email = ?,
      google_customer = ?,
      google_sync_enabled = ?,
      google_setup_completed = ?,
      google_credential_version = google_credential_version + ?,
      sync_new_google_users_by_default = ?,
      microsoft_tenant_id = ?,
      microsoft_client_id = ?,
      microsoft_client_secret_encrypted = COALESCE(?, microsoft_client_secret_encrypted),
      microsoft_credential_version = microsoft_credential_version + ?,
      microsoft_test_user_email = ?,
      microsoft_sync_enabled = ?,
      microsoft_setup_completed = ?,
      sync_new_microsoft_users_by_default = ?,
      microsoft_sync_policy_json = ?,
      microsoft_consent_granted_at = ?,
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
      Number(schoolboxSetupCompleted),
      Number(schoolboxConnectionChanged),
      serviceAccountEncrypted,
      adminEmail || null,
      customer || "my_customer",
      Number(googleEnabled),
      Number(googleSetupCompleted),
      Number(googleConnectionChanged),
      Number(syncNewGoogleUsersByDefault),
      microsoftTenantId || null,
      microsoftClientId || null,
      microsoftClientSecretEncrypted,
      Number(microsoftConnectionChanged),
      microsoftTestUserEmail || null,
      Number(microsoftEnabled),
      Number(microsoftSetupCompleted),
      Number(syncNewMicrosoftUsersByDefault),
      JSON.stringify(microsoftSyncPolicy),
      effectiveMicrosoftConsent ?? null,
      timezone,
      clampInteger(input.pastDays, current.pastDays, 0, 365),
      clampInteger(input.futureDays, current.futureDays, 1, 730),
      clampInteger(input.concurrency, current.concurrency, 1, 10),
      clampInteger(input.discoveryTimeoutSeconds, current.discoveryTimeoutSeconds, 30, 900),
      clampInteger(input.userSyncTimeoutSeconds, current.userSyncTimeoutSeconds, 30, 1800),
      clampInteger(input.runTimeoutMinutes, current.runTimeoutMinutes, 5, 240),
      clampInteger(input.syncIntervalMinutes, current.syncIntervalMinutes, 15, 1440),
      Number(syncNewGoogleUsersByDefault),
      JSON.stringify(syncPolicy),
      Number(schedulerEnabled),
      Number(setupCompleted),
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

async function schoolboxUserIdForTarget(targetUserId: string, target: TargetProvider): Promise<number | null> {
  await ensureSchema();
  return db().prepare(`SELECT schoolbox_user_id AS schoolboxUserId FROM target_user_mappings
    WHERE target = ? AND target_user_id = ? AND directory_active = 1`)
    .bind(target, targetUserId)
    .first<{ schoolboxUserId: number | null }>()?.schoolboxUserId ?? null;
}

export async function getUserEventExclusions(
  targetUserId: string,
  target: TargetProvider = "google",
): Promise<UserEventExclusions> {
  await ensureSchema();
  const schoolboxUserId = await schoolboxUserIdForTarget(targetUserId, target);
  if (schoolboxUserId === null) return normalizeUserEventExclusions(null);
  const row = db().prepare(`SELECT excluded_categories_json AS categoriesJson,
    excluded_event_types_json AS eventTypesJson, updated_at AS updatedAt, updated_by AS updatedBy
    FROM schoolbox_user_exclusions WHERE schoolbox_user_id = ?`)
    .bind(schoolboxUserId)
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
  targetUserId: string,
  input: UserEventExclusionsInput,
  actor: string,
  target: TargetProvider = "google",
): Promise<UserEventExclusions> {
  await ensureSchema();
  const id = targetUserId.trim();
  if (!id || id.length > 200) throw new HttpError(400, "Choose a valid user");
  const binding = db();
  const user = binding.prepare(`SELECT schoolbox_user_id AS schoolboxUserId FROM target_user_mappings
    WHERE target = ? AND target_user_id = ? AND directory_active = 1`)
    .bind(target, id)
    .first<{ schoolboxUserId: number | null }>();
  if (!user) throw new HttpError(404, "User not found");
  if (user.schoolboxUserId === null) {
    throw new HttpError(409, "Event exclusions can only be configured after a Schoolbox identity is matched");
  }

  const now = new Date().toISOString();
  const normalized = normalizeUserEventExclusions(input, { updatedAt: now, updatedBy: actor });
  binding.transaction(() => {
    if (normalized.categories.length === 0 && normalized.eventTypes.length === 0) {
      binding.prepare("DELETE FROM schoolbox_user_exclusions WHERE schoolbox_user_id = ?").bind(user.schoolboxUserId).run();
    } else {
      binding.prepare(`INSERT INTO schoolbox_user_exclusions
        (schoolbox_user_id, excluded_categories_json, excluded_event_types_json, updated_at, updated_by)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(schoolbox_user_id) DO UPDATE SET
          excluded_categories_json = excluded.excluded_categories_json,
          excluded_event_types_json = excluded.excluded_event_types_json,
          updated_at = excluded.updated_at,
          updated_by = excluded.updated_by`)
        .bind(user.schoolboxUserId, JSON.stringify(normalized.categories), JSON.stringify(normalized.eventTypes), now, actor)
        .run();
    }
    binding.prepare(`INSERT INTO audit_log (occurred_at, actor, action, detail)
      VALUES (?, ?, 'user.event_exclusions_updated', ?)`)
      .bind(
        now,
        actor,
        `${target}:${id}: ${normalized.categories.length} category exclusion(s), ${normalized.eventTypes.length} exact type exclusion(s)`,
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

export async function startRunTarget(runId: string, target: TargetProvider): Promise<RunTargetSummary> {
  await ensureSchema();
  const startedAt = new Date().toISOString();
  db().prepare(`INSERT INTO sync_run_targets
    (run_id, target, status, phase, phase_detail, started_at)
    VALUES (?, ?, 'running', 'discovery', 'Preparing target discovery.', ?)
    ON CONFLICT(run_id, target) DO UPDATE SET status='running', phase='discovery',
      phase_detail='Preparing target discovery.', started_at=excluded.started_at, completed_at=NULL,
      users_discovered=0, users_matched=0, users_selected=0, users_synced=0,
      events_created=0, events_updated=0, events_deleted=0, events_unchanged=0,
      errors=0, message=NULL`)
    .bind(runId, target, startedAt).run();
  return {
    runId, target, status: "running", phase: "discovery", phaseDetail: "Preparing target discovery.",
    startedAt, completedAt: null, usersDiscovered: 0, usersMatched: 0, usersSelected: 0,
    usersSynced: 0, eventsCreated: 0, eventsUpdated: 0, eventsDeleted: 0,
    eventsUnchanged: 0, errors: 0, message: null,
  };
}

export async function checkpointRunTarget(summary: RunTargetSummary, phase: string, detail: string): Promise<void> {
  await ensureSchema();
  summary.phase = phase;
  summary.phaseDetail = detail;
  const now = new Date().toISOString();
  const safeDetail = detail.slice(0, 2_000);
  const binding = db();
  binding.transaction(() => {
    binding.prepare(`UPDATE sync_run_targets SET phase=?, phase_detail=?, users_discovered=?,
      users_matched=?, users_selected=?, users_synced=?, events_created=?, events_updated=?,
      events_deleted=?, events_unchanged=?, errors=? WHERE run_id=? AND target=?`)
      .bind(phase, safeDetail, summary.usersDiscovered, summary.usersMatched,
        summary.usersSelected, summary.usersSynced, summary.eventsCreated, summary.eventsUpdated,
        summary.eventsDeleted, summary.eventsUnchanged, summary.errors, summary.runId, summary.target).run();
    const label = summary.target === "google" ? "Google Workspace" : "Microsoft 365";
    binding.prepare(`UPDATE sync_runs SET phase='target_sync', phase_detail=?, progress_at=?, heartbeat_at=?
      WHERE id=? AND status='running'`).bind(`${label}: ${safeDetail}`.slice(0, 2_000), now, now, summary.runId).run();
  });
}

export async function finishRunTarget(summary: RunTargetSummary): Promise<void> {
  await ensureSchema();
  summary.completedAt ??= new Date().toISOString();
  db().prepare(`UPDATE sync_run_targets SET status=?, phase=?, phase_detail=?, completed_at=?,
    users_discovered=?, users_matched=?, users_selected=?, users_synced=?, events_created=?,
    events_updated=?, events_deleted=?, events_unchanged=?, errors=?, message=?
    WHERE run_id=? AND target=?`)
    .bind(summary.status, summary.phase, summary.phaseDetail?.slice(0, 2_000) ?? null,
      summary.completedAt, summary.usersDiscovered, summary.usersMatched, summary.usersSelected,
      summary.usersSynced, summary.eventsCreated, summary.eventsUpdated, summary.eventsDeleted,
      summary.eventsUnchanged, summary.errors, summary.message?.slice(0, 2_000) ?? null,
      summary.runId, summary.target).run();
}

export async function listRunTargets(runId: string): Promise<RunTargetSummary[]> {
  await ensureSchema();
  return db().prepare(`SELECT run_id AS runId, target, status, phase, phase_detail AS phaseDetail,
    started_at AS startedAt, completed_at AS completedAt, users_discovered AS usersDiscovered,
    users_matched AS usersMatched, users_selected AS usersSelected, users_synced AS usersSynced,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, errors, message FROM sync_run_targets
    WHERE run_id = ? ORDER BY target`).bind(runId).all<RunTargetSummary>().results;
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
  const now = new Date().toISOString();
  const cutoff = new Date(Date.now() - maxAgeMinutes * 60_000).toISOString();
  const binding = db();
  return binding.transaction(() => {
    const stale = `SELECT id FROM sync_runs WHERE status = 'running' AND COALESCE(heartbeat_at, started_at) <= ?`;
    binding.prepare(`UPDATE sync_run_targets SET status = 'failed', phase = 'failed', completed_at = ?,
      errors = errors + 1, message = 'Parent run heartbeat stopped before completion.'
      WHERE status = 'running' AND run_id IN (${stale})`).bind(now, cutoff).run();
    binding.prepare(`UPDATE sync_run_target_users SET status = 'failed', stage = 'interrupted', completed_at = ?,
      error_message = 'Run heartbeat stopped before this account completed.'
      WHERE status = 'running' AND run_id IN (${stale})`).bind(now, cutoff).run();
    // Retain repair support for diagnostics created by pre-provider releases.
    binding.prepare(`UPDATE sync_run_users SET status = 'failed', stage = 'interrupted', completed_at = ?,
      error_message = 'Run heartbeat stopped before this account completed.'
      WHERE status = 'running' AND run_id IN (${stale})`).bind(now, cutoff).run();
    const result = binding.prepare(`UPDATE sync_runs SET status = 'failed', phase = 'failed',
      phase_detail = 'Run heartbeat stopped before completion.', completed_at = ?, errors = errors + 1,
      message = 'Run was interrupted by a server restart or exceeded the maximum runtime.'
      WHERE status = 'running' AND COALESCE(heartbeat_at, started_at) <= ?`).bind(now, cutoff).run();
    return Number(result.changes);
  });
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
  db().prepare(`DELETE FROM sync_run_target_events WHERE run_id IN (${expiredRuns})`).run();
  db().prepare(`DELETE FROM sync_run_target_users WHERE run_id IN (${expiredRuns})`).run();
  db().prepare(`DELETE FROM sync_run_targets WHERE run_id IN (${expiredRuns})`).run();
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

type RunUserStartInput = Pick<RunUserDiagnostic, "runId" | "displayName" | "schoolboxUserId" | "schoolboxEmail"> & {
  target?: TargetProvider;
  targetUserId?: string;
  targetEmail?: string;
  googleUserId?: string;
  googleEmail?: string;
};

export async function startRunUserDiagnostic(input: RunUserStartInput): Promise<void> {
  await ensureSchema();
  const target = input.target ?? "google";
  const targetUserId = input.targetUserId ?? input.googleUserId;
  const targetEmail = input.targetEmail ?? input.googleEmail;
  if (!targetUserId || !targetEmail) throw new HttpError(400, "Target user details are required");
  const startedAt = new Date().toISOString();
  db().prepare(`INSERT INTO sync_run_target_users
    (run_id, target, target_user_id, target_email, display_name, schoolbox_user_id, schoolbox_email,
     status, stage, started_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, 'running', 'starting', ?)
    ON CONFLICT(run_id, target, target_user_id) DO UPDATE SET
      target_email=excluded.target_email, display_name=excluded.display_name,
      schoolbox_user_id=excluded.schoolbox_user_id, schoolbox_email=excluded.schoolbox_email,
      status='running', stage='starting', started_at=excluded.started_at, completed_at=NULL,
      events_found=0, events_included=0, events_created=0, events_updated=0,
      events_deleted=0, events_unchanged=0, managed_events_after=0, error_message=NULL`)
    .bind(
      input.runId,
      target,
      targetUserId,
      targetEmail,
      input.displayName,
      input.schoolboxUserId,
      input.schoolboxEmail,
      startedAt,
    ).run();
}

export async function finishRunUserDiagnostic(input: Pick<RunUserDiagnostic,
  "runId" | "status" | "stage" | "eventsFound" | "eventsIncluded" |
  "eventsCreated" | "eventsUpdated" | "eventsDeleted" | "eventsUnchanged" |
  "managedEventsAfter" | "errorMessage"> & {
    target?: TargetProvider; targetUserId?: string; googleUserId?: string;
  }): Promise<void> {
  await ensureSchema();
  const target = input.target ?? "google";
  const targetUserId = input.targetUserId ?? input.googleUserId;
  if (!targetUserId) throw new HttpError(400, "Target user is required");
  db().prepare(`UPDATE sync_run_target_users SET status = ?, stage = ?, completed_at = ?,
    events_found = ?, events_included = ?, events_created = ?, events_updated = ?,
    events_deleted = ?, events_unchanged = ?, managed_events_after = ?, error_message = ?
    WHERE run_id = ? AND target = ? AND target_user_id = ?`)
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
      target,
      targetUserId,
    ).run();
}

export async function recordRunEventDiagnostic(input: Omit<RunEventDiagnostic,
  "recordedAt" | "target" | "targetUserId" | "targetEventId" | "googleUserId" | "googleEventId"> & {
    target?: TargetProvider; targetUserId?: string; targetEventId?: string | null;
    googleUserId?: string; googleEventId?: string | null;
  }): Promise<void> {
  await ensureSchema();
  const target = input.target ?? "google";
  const targetUserId = input.targetUserId ?? input.googleUserId;
  const targetEventId = input.targetEventId ?? input.googleEventId;
  if (!targetUserId) throw new HttpError(400, "Target user is required");
  db().prepare(`INSERT INTO sync_run_target_events
    (run_id, target, target_user_id, source_key, title, description, location, author, event_type,
     category, source_start, source_end, all_day, source_url, target_event_id, calendar_id,
     destination_id, action, detail, error_message, recorded_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(run_id, target, target_user_id, source_key) DO UPDATE SET
      title=excluded.title, description=excluded.description, location=excluded.location,
      author=excluded.author, event_type=excluded.event_type, category=excluded.category,
      source_start=excluded.source_start, source_end=excluded.source_end, all_day=excluded.all_day,
      source_url=excluded.source_url, target_event_id=excluded.target_event_id,
      calendar_id=excluded.calendar_id, destination_id=excluded.destination_id,
      action=excluded.action, detail=excluded.detail, error_message=excluded.error_message,
      recorded_at=excluded.recorded_at`)
    .bind(
      input.runId,
      target,
      targetUserId,
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
      targetEventId ?? null,
      input.calendarId,
      input.destinationId,
      input.action,
      input.detail?.slice(0, 2_000) ?? null,
      input.errorMessage?.slice(0, 4_000) ?? null,
      new Date().toISOString(),
    ).run();
}

export async function listRunUserDiagnostics(runId: string, target?: TargetProvider): Promise<RunUserDiagnostic[]> {
  await ensureSchema();
  const statement = db().prepare(`SELECT target, target_user_id AS targetUserId,
    target_email AS targetEmail, target_user_id AS googleUserId, target_email AS googleEmail,
    run_id AS runId, display_name AS displayName, schoolbox_user_id AS schoolboxUserId,
    schoolbox_email AS schoolboxEmail, status, stage, started_at AS startedAt,
    completed_at AS completedAt, events_found AS eventsFound, events_included AS eventsIncluded,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, managed_events_after AS managedEventsAfter,
    error_message AS errorMessage FROM sync_run_target_users WHERE run_id = ?${target ? " AND target = ?" : ""}
    ORDER BY CASE status WHEN 'failed' THEN 0 WHEN 'running' THEN 1 ELSE 2 END,
      target, target_email COLLATE NOCASE`);
  return (target ? statement.bind(runId, target) : statement.bind(runId)).all<RunUserDiagnostic>().results;
}

export async function listRunEventDiagnostics(
  runId: string,
  targetUserId: string,
  options: { limit?: number; offset?: number } = {},
  target: TargetProvider = "google",
): Promise<{ events: RunEventDiagnostic[]; total: number }> {
  await ensureSchema();
  const limit = Math.max(1, Math.min(options.limit ?? 100, 250));
  const offset = Math.max(0, options.offset ?? 0);
  const binding = db();
  const total = binding.prepare(`SELECT COUNT(*) AS count FROM sync_run_target_events
    WHERE run_id = ? AND target = ? AND target_user_id = ?`).bind(runId, target, targetUserId).first<{ count: number }>()?.count ?? 0;
  const events = binding.prepare(`SELECT target, target_user_id AS targetUserId,
    target_event_id AS targetEventId, target_user_id AS googleUserId,
    target_event_id AS googleEventId, run_id AS runId,
    source_key AS sourceKey, title, description, location, author, event_type AS eventType,
    category, source_start AS sourceStart, source_end AS sourceEnd, all_day AS allDay,
    source_url AS sourceUrl, calendar_id AS calendarId,
    destination_id AS destinationId, action, detail, error_message AS errorMessage,
    recorded_at AS recordedAt FROM sync_run_target_events
    WHERE run_id = ? AND target = ? AND target_user_id = ?
    ORDER BY CASE action WHEN 'failed' THEN 0 WHEN 'created' THEN 1 WHEN 'updated' THEN 2
      WHEN 'deleted' THEN 3 ELSE 4 END, source_start, title COLLATE NOCASE LIMIT ? OFFSET ?`)
    .bind(runId, target, targetUserId, limit, offset).all<RunEventDiagnostic>().results;
  for (const event of events) event.allDay = Boolean(event.allDay);
  return { events, total };
}

export async function listUserRunDiagnostics(
  targetUserId: string,
  limit = 20,
  target: TargetProvider = "google",
): Promise<RunUserDiagnostic[]> {
  await ensureSchema();
  return db().prepare(`SELECT target, target_user_id AS targetUserId, target_email AS targetEmail,
    target_user_id AS googleUserId, target_email AS googleEmail, run_id AS runId,
    display_name AS displayName, schoolbox_user_id AS schoolboxUserId,
    schoolbox_email AS schoolboxEmail, status, stage, started_at AS startedAt,
    completed_at AS completedAt, events_found AS eventsFound, events_included AS eventsIncluded,
    events_created AS eventsCreated, events_updated AS eventsUpdated, events_deleted AS eventsDeleted,
    events_unchanged AS eventsUnchanged, managed_events_after AS managedEventsAfter,
    error_message AS errorMessage FROM sync_run_target_users WHERE target = ? AND target_user_id = ?
    ORDER BY started_at DESC LIMIT ?`)
    .bind(target, targetUserId, Math.max(1, Math.min(limit, 100))).all<RunUserDiagnostic>().results;
}

export type UserMappingWrite = Pick<UserMapping,
  "schoolboxUserId" | "schoolboxEmail" | "displayName" | "role" | "status" |
  "lastSyncAt" | "lastError" | "eventCount" | "updatedAt"> & {
    target?: TargetProvider;
    targetUserId?: string;
    targetEmail?: string;
    googleUserId?: string;
    googleEmail?: string;
  };

function targetIdentity(
  mapping: Pick<UserMappingWrite, "target" | "targetUserId" | "targetEmail" | "googleUserId" | "googleEmail">,
  fallback: TargetProvider = "google",
): { target: TargetProvider; id: string; email: string } {
  const target = mapping.target ?? fallback;
  const id = (mapping.targetUserId ?? mapping.googleUserId ?? "").trim();
  const email = (mapping.targetEmail ?? mapping.googleEmail ?? "").trim().toLowerCase();
  if (!id || !email) throw new HttpError(400, "Target user ID and email are required");
  return { target, id, email };
}

export async function upsertUserMapping(mapping: UserMappingWrite, fallbackTarget: TargetProvider = "google"): Promise<void> {
  await ensureSchema();
  const identity = targetIdentity(mapping, fallbackTarget);
  await db()
    .prepare(`INSERT INTO target_user_mappings
      (target, target_user_id, target_email, schoolbox_user_id, schoolbox_email, display_name, role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 1, ?)
      ON CONFLICT(target, target_user_id) DO UPDATE SET target_email=excluded.target_email, schoolbox_user_id=excluded.schoolbox_user_id,
      schoolbox_email=excluded.schoolbox_email, display_name=excluded.display_name, role=excluded.role, status=excluded.status,
      last_sync_at=excluded.last_sync_at, last_error=excluded.last_error, event_count=excluded.event_count,
      directory_active=1, updated_at=excluded.updated_at`)
    .bind(
      identity.target,
      identity.id,
      identity.email,
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
  fallbackTarget: TargetProvider = "google",
): Promise<Map<string, boolean>> {
  await ensureSchema();
  const binding = db();
  const upsert = binding.prepare(`INSERT INTO target_user_mappings
    (target, target_user_id, target_email, schoolbox_user_id, schoolbox_email, display_name, role, status, last_sync_at, last_error, event_count, sync_enabled, directory_active, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
    ON CONFLICT(target, target_user_id) DO UPDATE SET
      target_email=excluded.target_email,
      schoolbox_user_id=excluded.schoolbox_user_id,
      schoolbox_email=excluded.schoolbox_email,
      display_name=excluded.display_name,
      role=excluded.role,
      status=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN 'unmatched'
        WHEN target_user_mappings.status = 'unmatched' THEN 'pending'
        ELSE target_user_mappings.status
      END,
      last_sync_at=target_user_mappings.last_sync_at,
      last_error=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN excluded.last_error
        WHEN target_user_mappings.status = 'unmatched' THEN NULL
        ELSE target_user_mappings.last_error
      END,
      event_count=target_user_mappings.event_count,
      sync_enabled=CASE
        WHEN excluded.schoolbox_user_id IS NULL THEN 0
        ELSE target_user_mappings.sync_enabled
      END,
      directory_active=1,
      updated_at=excluded.updated_at`);
  const selection = binding.prepare("SELECT sync_enabled FROM target_user_mappings WHERE target = ? AND target_user_id = ?");

  return binding.transaction(() => {
    const result = new Map<string, boolean>();
    const discoveryTime = discoveries[0]?.updatedAt ?? new Date().toISOString();
    // Email addresses are mutable. Retiring the prior snapshot first makes
    // swaps and reassignment order-independent while the stable target ID keeps
    // each user's selection and event mappings attached to the correct row.
    const target = discoveries[0]?.target ?? fallbackTarget;
    binding.prepare("UPDATE target_user_mappings SET directory_active = 0, updated_at = ? WHERE target = ? AND directory_active = 1")
      .bind(discoveryTime, target).run();
    for (const mapping of discoveries) {
      const identity = targetIdentity(mapping, fallbackTarget);
      if (identity.target !== target) throw new HttpError(400, "A directory discovery must contain one target only");
      upsert.bind(
        identity.target,
        identity.id,
        identity.email,
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
      const row = selection.bind(identity.target, identity.id).first<{ sync_enabled: number }>();
      result.set(identity.id, Boolean(row?.sync_enabled));
    }
    return result;
  });
}

export async function setUsersSyncEnabled(
  ids: string[],
  enabled: boolean,
  actor: string,
  target: TargetProvider = "google",
): Promise<number> {
  await ensureSchema();
  const uniqueIds = [...new Set(ids.map((id) => id.trim()).filter(Boolean))];
  if (uniqueIds.length === 0) throw new HttpError(400, "Choose at least one user");
  if (uniqueIds.length > 25_000) throw new HttpError(400, "Update no more than 25,000 users at a time");

  const binding = db();
  const exists = binding.prepare("SELECT target_user_id, schoolbox_user_id FROM target_user_mappings WHERE target = ? AND target_user_id = ? AND directory_active = 1");
  const update = binding.prepare("UPDATE target_user_mappings SET sync_enabled = ?, updated_at = ? WHERE target = ? AND target_user_id = ? AND directory_active = 1");
  const audit = binding.prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, 'users.sync_selection_updated', ?)");
  const now = new Date().toISOString();
  const updated = binding.transaction(() => {
    for (const id of uniqueIds) {
      const user = exists.bind(target, id).first<{ target_user_id: string; schoolbox_user_id: string | null }>();
      if (!user) throw new HttpError(404, "One or more users are no longer available");
      if (enabled && user.schoolbox_user_id === null) {
        throw new HttpError(409, "Unmatched users cannot be enabled until a Schoolbox identity is discovered");
      }
    }
    let changes = 0;
    for (const id of uniqueIds) {
      changes += Number(update.bind(Number(enabled), now, target, id).run().changes);
    }
    audit.bind(now, actor, `${target}: ${uniqueIds.length} user selection(s) set to ${enabled ? "enabled" : "paused"}`).run();
    return changes;
  });
  return updated;
}

export async function listUserMappings(
  limit?: number,
  includeInactive = false,
  target: TargetProvider = "google",
): Promise<UserMapping[]> {
  await ensureSchema();
  const statement = db()
    .prepare(`SELECT u.target, u.target_user_id AS targetUserId, u.target_email AS targetEmail,
      u.target_user_id AS googleUserId, u.target_email AS googleEmail, u.schoolbox_user_id AS schoolboxUserId,
      u.schoolbox_email AS schoolboxEmail, u.display_name AS displayName, u.role, u.status, u.last_sync_at AS lastSyncAt,
      u.last_error AS lastError,
      (SELECT COUNT(*) FROM target_event_mappings e WHERE e.target = u.target AND e.target_user_id = u.target_user_id) AS eventCount,
      (SELECT COUNT(*) FROM target_calendar_targets c WHERE c.target = u.target AND c.target_user_id = u.target_user_id) AS calendarCount,
      EXISTS(SELECT 1 FROM schoolbox_user_exclusions x WHERE x.schoolbox_user_id = u.schoolbox_user_id) AS hasCustomExclusions,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM target_user_mappings u WHERE u.target = ?${includeInactive ? "" : " AND u.directory_active = 1"}
      ORDER BY u.target_email${limit === undefined ? "" : " LIMIT ?"}`);
  const result = limit === undefined
    ? statement.bind(target).all<UserMapping>()
    : statement.bind(target, Math.max(1, Math.min(limit, 5000))).all<UserMapping>();
  for (const mapping of result.results) {
    mapping.hasCustomExclusions = Boolean(mapping.hasCustomExclusions);
    mapping.syncEnabled = Boolean(mapping.syncEnabled);
    mapping.directoryActive = Boolean(mapping.directoryActive);
  }
  return result.results;
}

export async function getUserMapping(targetUserId: string, target: TargetProvider = "google"): Promise<UserMapping | null> {
  await ensureSchema();
  const mapping = db()
    .prepare(`SELECT u.target, u.target_user_id AS targetUserId, u.target_email AS targetEmail,
      u.target_user_id AS googleUserId, u.target_email AS googleEmail, u.schoolbox_user_id AS schoolboxUserId,
      u.schoolbox_email AS schoolboxEmail, u.display_name AS displayName, u.role, u.status, u.last_sync_at AS lastSyncAt,
      u.last_error AS lastError,
      (SELECT COUNT(*) FROM target_event_mappings e WHERE e.target = u.target AND e.target_user_id = u.target_user_id) AS eventCount,
      (SELECT COUNT(*) FROM target_calendar_targets c WHERE c.target = u.target AND c.target_user_id = u.target_user_id) AS calendarCount,
      EXISTS(SELECT 1 FROM schoolbox_user_exclusions x WHERE x.schoolbox_user_id = u.schoolbox_user_id) AS hasCustomExclusions,
      u.sync_enabled AS syncEnabled, u.directory_active AS directoryActive, u.updated_at AS updatedAt
      FROM target_user_mappings u WHERE u.target = ? AND u.target_user_id = ? AND u.directory_active = 1`)
    .bind(target, targetUserId)
    .first<UserMapping>();
  if (mapping) {
    mapping.hasCustomExclusions = Boolean(mapping.hasCustomExclusions);
    mapping.syncEnabled = Boolean(mapping.syncEnabled);
    mapping.directoryActive = Boolean(mapping.directoryActive);
  }
  return mapping;
}

export async function getEventMappings(
  targetUserId: string,
  target: TargetProvider = "google",
): Promise<EventMapping[]> {
  await ensureSchema();
  const result = await db()
    .prepare(`SELECT target, target_user_id AS targetUserId, target_event_id AS targetEventId,
      target_user_id AS googleUserId, target_event_id AS googleEventId, source_key AS sourceKey,
      calendar_id AS calendarId, source_hash AS sourceHash, source_start AS sourceStart,
      source_end AS sourceEnd, last_seen_run_id AS lastSeenRunId, created_at AS createdAt,
      updated_at AS updatedAt, title, description, location, author, event_type AS eventType,
      category, all_day AS allDay, source_url AS sourceUrl, destination_id AS destinationId
      FROM target_event_mappings WHERE target = ? AND target_user_id = ?
      ORDER BY source_start, title COLLATE NOCASE`)
    .bind(target, targetUserId)
    .all<EventMapping>();
  for (const event of result.results) event.allDay = Boolean(event.allDay);
  return result.results;
}

type EventMappingDiagnosticFields = "title" | "description" | "location" | "author" |
  "eventType" | "category" | "allDay" | "sourceUrl" | "destinationId";

export async function upsertEventMapping(
  mapping: Pick<EventMapping, "sourceKey" | "sourceHash" | "sourceStart" | "sourceEnd" |
    "lastSeenRunId" | "createdAt" | "updatedAt"> & {
      target?: TargetProvider;
      targetUserId?: string;
      targetEventId?: string;
      googleUserId?: string;
      googleEventId?: string;
      calendarId?: string;
    } & Partial<Pick<EventMapping, EventMappingDiagnosticFields>>,
): Promise<void> {
  await ensureSchema();
  const target = mapping.target ?? "google";
  const targetUserId = mapping.targetUserId ?? mapping.googleUserId;
  const targetEventId = mapping.targetEventId ?? mapping.googleEventId;
  if (!targetUserId || !targetEventId) throw new HttpError(400, "Target event identity is required");
  await db()
    .prepare(`INSERT INTO target_event_mappings
      (target, target_user_id, source_key, target_event_id, calendar_id, source_hash, source_start, source_end,
       last_seen_run_id, created_at, updated_at, title, description, location, author, event_type,
       category, all_day, source_url, destination_id)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(target, target_user_id, source_key) DO UPDATE SET target_event_id=excluded.target_event_id,
      calendar_id=excluded.calendar_id, source_hash=excluded.source_hash, source_start=excluded.source_start, source_end=excluded.source_end,
      last_seen_run_id=excluded.last_seen_run_id, updated_at=excluded.updated_at,
      title=excluded.title, description=excluded.description, location=excluded.location,
      author=excluded.author, event_type=excluded.event_type, category=excluded.category,
      all_day=excluded.all_day, source_url=excluded.source_url, destination_id=excluded.destination_id`)
    .bind(
      target,
      targetUserId,
      mapping.sourceKey,
      targetEventId,
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
  targetUserId: string,
  destinationId: string,
  target: TargetProvider = "google",
): Promise<UserCalendarTarget | null> {
  await ensureSchema();
  return db().prepare(`SELECT target, target_user_id AS targetUserId, target_calendar_id AS targetCalendarId,
    target_user_id AS googleUserId, target_calendar_id AS googleCalendarId, destination_id AS destinationId,
    summary, description, time_zone AS timeZone,
    created_at AS createdAt, updated_at AS updatedAt
    FROM target_calendar_targets WHERE target = ? AND target_user_id = ? AND destination_id = ?`)
    .bind(target, targetUserId, destinationId)
    .first<UserCalendarTarget>();
}

export async function listUserCalendarTargets(
  targetUserId: string,
  target: TargetProvider = "google",
): Promise<UserCalendarTarget[]> {
  await ensureSchema();
  return db().prepare(`SELECT target, target_user_id AS targetUserId, target_calendar_id AS targetCalendarId,
    target_user_id AS googleUserId, target_calendar_id AS googleCalendarId, destination_id AS destinationId,
    summary, description, time_zone AS timeZone,
    created_at AS createdAt, updated_at AS updatedAt
    FROM target_calendar_targets WHERE target = ? AND target_user_id = ? ORDER BY destination_id`)
    .bind(target, targetUserId)
    .all<UserCalendarTarget>().results;
}

export async function listCalendarDestinationUsage(target: TargetProvider = "google"): Promise<CalendarDestinationUsage[]> {
  await ensureSchema();
  return db().prepare(`SELECT t.target, t.destination_id AS destinationId, MAX(t.summary) AS summary,
    COUNT(DISTINCT t.target_user_id) AS calendarCount, COUNT(e.source_key) AS eventCount
    FROM target_calendar_targets t
    LEFT JOIN target_event_mappings e ON e.target = t.target AND e.target_user_id = t.target_user_id
      AND e.calendar_id = t.target_calendar_id
    WHERE t.target = ? GROUP BY t.target, t.destination_id ORDER BY MAX(t.summary) COLLATE NOCASE`)
    .bind(target)
    .all<CalendarDestinationUsage>().results;
}

export async function listCalendarTargetsForDestination(
  destinationId: string,
  target: TargetProvider = "google",
): Promise<UserCalendarTargetWithEmail[]> {
  await ensureSchema();
  return db().prepare(`SELECT t.target, t.target_user_id AS targetUserId, u.target_email AS targetEmail,
    t.target_calendar_id AS targetCalendarId, t.target_user_id AS googleUserId,
    u.target_email AS googleEmail, t.destination_id AS destinationId,
    t.target_calendar_id AS googleCalendarId,
    t.summary, t.description, t.time_zone AS timeZone, t.created_at AS createdAt, t.updated_at AS updatedAt
    FROM target_calendar_targets t
    JOIN target_user_mappings u ON u.target = t.target AND u.target_user_id = t.target_user_id
    WHERE t.target = ? AND t.destination_id = ? ORDER BY u.target_email COLLATE NOCASE`)
    .bind(target, destinationId)
    .all<UserCalendarTargetWithEmail>().results;
}

export async function deleteCalendarTargetRecords(
  targetUserId: string,
  destinationId: string,
  targetCalendarId: string,
  provider: TargetProvider = "google",
): Promise<number> {
  await ensureSchema();
  const binding = db();
  return binding.transaction(() => {
    const target = binding.prepare(`SELECT 1 AS found FROM target_calendar_targets
      WHERE target = ? AND target_user_id = ? AND destination_id = ? AND target_calendar_id = ?`)
      .bind(provider, targetUserId, destinationId, targetCalendarId)
      .first<{ found: number }>();
    if (!target) return 0;
    const removedEvents = Number(binding.prepare(`DELETE FROM target_event_mappings
      WHERE target = ? AND target_user_id = ? AND calendar_id = ?`)
      .bind(provider, targetUserId, targetCalendarId).run().changes);
    binding.prepare(`DELETE FROM target_calendar_targets
      WHERE target = ? AND target_user_id = ? AND destination_id = ? AND target_calendar_id = ?`)
      .bind(provider, targetUserId, destinationId, targetCalendarId).run();
    return removedEvents;
  });
}

export async function upsertUserCalendarTarget(input: Omit<UserCalendarTarget,
  "target" | "targetUserId" | "targetCalendarId"> & {
    target?: TargetProvider; targetUserId?: string; targetCalendarId?: string;
  }): Promise<void> {
  await ensureSchema();
  const target = input.target ?? "google";
  const targetUserId = input.targetUserId ?? input.googleUserId;
  const targetCalendarId = input.targetCalendarId ?? input.googleCalendarId;
  if (!targetUserId || !targetCalendarId) throw new HttpError(400, "Target calendar identity is required");
  await db().prepare(`INSERT INTO target_calendar_targets
    (target, target_user_id, destination_id, target_calendar_id, summary, description, time_zone, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(target, target_user_id, destination_id) DO UPDATE SET
      target_calendar_id=excluded.target_calendar_id, summary=excluded.summary,
      description=excluded.description, time_zone=excluded.time_zone, updated_at=excluded.updated_at`)
    .bind(
      target,
      targetUserId,
      input.destinationId,
      targetCalendarId,
      input.summary,
      input.description,
      input.timeZone,
      input.createdAt,
      input.updatedAt,
    )
    .run();
}

export async function touchEventMapping(
  targetUserId: string,
  sourceKey: string,
  runId: string,
  diagnostic: Partial<Pick<EventMapping, EventMappingDiagnosticFields>> = {},
  target: TargetProvider = "google",
): Promise<void> {
  await ensureSchema();
  await db()
    .prepare(`UPDATE target_event_mappings SET last_seen_run_id = ?, updated_at = ?,
      title = COALESCE(?, title), description = COALESCE(?, description),
      location = COALESCE(?, location), author = COALESCE(?, author),
      event_type = COALESCE(?, event_type), category = COALESCE(?, category),
      all_day = ?, source_url = COALESCE(?, source_url),
      destination_id = COALESCE(?, destination_id)
      WHERE target = ? AND target_user_id = ? AND source_key = ?`)
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
      target,
      targetUserId,
      sourceKey,
    )
    .run();
}

export async function deleteEventMapping(
  targetUserId: string,
  sourceKey: string,
  target: TargetProvider = "google",
): Promise<void> {
  await ensureSchema();
  await db().prepare("DELETE FROM target_event_mappings WHERE target = ? AND target_user_id = ? AND source_key = ?")
    .bind(target, targetUserId, sourceKey).run();
}

export async function recordManagedEventCleanup(options: {
  target?: TargetProvider;
  targetUserId?: string;
  googleUserId?: string;
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
  const target = options.target ?? "google";
  const targetUserId = options.targetUserId ?? options.googleUserId;
  if (!targetUserId) throw new HttpError(400, "Target user is required");
  const binding = db();
  binding.transaction(() => {
    const result = binding.prepare(`UPDATE target_user_mappings SET sync_enabled = 0, event_count = ?,
      status = CASE WHEN ? IS NOT NULL THEN 'error' WHEN schoolbox_user_id IS NULL THEN 'unmatched' ELSE 'pending' END,
      last_error = ?, updated_at = ? WHERE target = ? AND target_user_id = ? AND directory_active = 1`)
      .bind(options.remaining, error, error, now, target, targetUserId)
      .run();
    if (Number(result.changes) !== 1) throw new HttpError(404, "This user is no longer available");
    binding.prepare("INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, ?, 'users.managed_events_cleanup', ?)")
      .bind(
        now,
        options.actor,
        `${target}: ${options.deleted} managed event(s) deleted, ${options.alreadyMissing} already absent, ${options.remaining} remaining; ${options.calendarsDeleted ?? 0} managed calendar(s) deleted, ${options.calendarsAlreadyMissing ?? 0} already absent, ${options.calendarsRemaining ?? 0} remaining`,
      )
      .run();
  });
}

export async function statusSnapshot(): Promise<{
  configured: boolean;
  config: AppConfig;
  lastRun: RunSummary | null;
  counts: { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number };
  targetCounts: Record<TargetProvider, { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number }>;
}> {
  const config = await getConfig(false);
  const countsFor = async (target: TargetProvider) => {
    const [users, events] = await Promise.all([
      db().prepare(`SELECT SUM(CASE WHEN directory_active = 1 THEN 1 ELSE 0 END) AS users,
      SUM(CASE WHEN directory_active = 1 AND schoolbox_user_id IS NOT NULL AND sync_enabled = 1 THEN 1 ELSE 0 END) AS enabled,
      SUM(CASE WHEN directory_active = 1 AND schoolbox_user_id IS NOT NULL AND sync_enabled = 0 THEN 1 ELSE 0 END) AS disabled,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'synced' THEN 1 ELSE 0 END) AS healthy,
      SUM(CASE WHEN directory_active = 1 AND sync_enabled = 1 AND status = 'error' THEN 1 ELSE 0 END) AS errors,
      SUM(CASE WHEN directory_active = 1 AND status = 'unmatched' THEN 1 ELSE 0 END) AS unmatched
      FROM target_user_mappings WHERE target = ?`).bind(target)
        .first<{ users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number }>(),
      db().prepare(`SELECT COUNT(*) AS count FROM target_event_mappings e
        JOIN target_user_mappings u ON u.target = e.target AND u.target_user_id = e.target_user_id
        WHERE e.target = ? AND u.directory_active = 1`).bind(target).first<{ count: number }>(),
    ]);
    return {
      users: users?.users ?? 0,
      enabled: users?.enabled ?? 0,
      disabled: users?.disabled ?? 0,
      healthy: users?.healthy ?? 0,
      errors: users?.errors ?? 0,
      unmatched: users?.unmatched ?? 0,
      events: events?.count ?? 0,
    };
  };
  const [runs, google, microsoft] = await Promise.all([listRuns(1), countsFor("google"), countsFor("microsoft")]);
  const targetCounts = { google, microsoft };
  const aggregate = TARGET_PROVIDERS.reduce((result, target) => {
    for (const key of Object.keys(result) as Array<keyof typeof result>) result[key] += targetCounts[target][key];
    return result;
  }, { users: 0, enabled: 0, disabled: 0, healthy: 0, errors: 0, unmatched: 0, events: 0 });
  return {
    configured: config.setupCompleted,
    config,
    lastRun: runs[0] ?? null,
    counts: aggregate,
    targetCounts,
  };
}
