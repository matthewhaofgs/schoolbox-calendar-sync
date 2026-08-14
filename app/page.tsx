"use client";

import { FormEvent, useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  DEFAULT_SYNC_POLICY,
  EVENT_CATEGORIES,
  eventTypeKey,
  normalizeEventTypeLabel,
  normalizeSyncPolicy,
  resolveGoogleEventRule,
  withoutManagedCalendarDestination,
  type EventCategory,
  type EventTypeFilterMode,
  type GoogleEventRuleOverride,
  type ManagedCalendarDefinition,
  type SyncPolicy,
  type UserEventExclusions,
} from "@/lib/policy";

type View = "dashboard" | "setup" | "people" | "runs" | "settings" | "access";
type TargetProvider = "google" | "microsoft";
type SetupTrack = "hub" | "schoolbox" | TargetProvider;
type SetupResume = {
  schoolboxTested: boolean;
  googleTested: boolean;
  microsoftTested: boolean;
  microsoftEnabled?: boolean;
  syncNewMicrosoftUsersByDefault?: boolean;
};
type Notice = { kind: "success" | "error" | "info"; message: string } | null;
type Permission = "view" | "operate" | "configure" | "manage_access";
type StaffRole = "admin" | "operator" | "viewer";

type AuthSession = {
  userId: string;
  actor: string;
  authType: "local" | "google";
  username: string | null;
  email: string | null;
  displayName: string;
  role: StaffRole;
  isOwner: boolean;
  permissions: Permission[];
  csrfToken: string;
  expiresAt: string;
  idleExpiresAt: string;
};

type AuthReadiness = { localAdministrator: boolean; googleSignInConfigured: boolean };

type StaffAccount = {
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

type OAuthSettings = {
  clientId: string;
  workspaceDomain: string;
  hasClientSecret: boolean;
  callbackUrl: string;
  configured: boolean;
};

type Person = {
  target: TargetProvider;
  id: string;
  name: string;
  schoolboxEmail: string;
  targetEmail: string;
  googleEmail: string;
  role: string;
  status: "Synced" | "Syncing" | "Pending" | "Unmatched" | "Error";
  hasCustomExclusions: boolean;
  syncEnabled: boolean;
  eventCount: number;
  calendarCount: number;
  lastSync: string;
};

type TargetCounts = { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number };

const EMPTY_TARGET_COUNTS: TargetCounts = { users: 0, enabled: 0, disabled: 0, healthy: 0, errors: 0, unmatched: 0, events: 0 };
const TARGET_LABELS: Record<TargetProvider, string> = { google: "Google Workspace", microsoft: "Microsoft 365" };
const TARGET_CALENDAR_LABELS: Record<TargetProvider, string> = { google: "Google Calendar", microsoft: "Outlook Calendar" };

type CalendarDestinationUsage = {
  destinationId: string;
  summary: string;
  calendarCount: number;
  eventCount: number;
};

type Run = {
  id: string;
  started: string;
  trigger: string;
  status: "Succeeded" | "Running" | "Warning" | "Failed";
  users: number;
  usersDiscovered: number;
  usersMatched: number;
  changes: number;
  duration: string;
  note: string;
  phase: string;
  phaseDetail: string;
  progressAt: string | null;
  created?: number;
  updated?: number;
  deleted?: number;
  errors?: number;
  targets?: RunTargetDiagnostic[];
};

type RunTargetDiagnostic = {
  runId?: string;
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

type RunUserDiagnostic = {
  target?: TargetProvider;
  targetUserId?: string;
  targetEmail?: string;
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

type DiagnosticEvent = {
  target?: TargetProvider;
  targetUserId?: string;
  targetEventId?: string | null;
  runId?: string;
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
  action?: string;
  detail?: string | null;
  errorMessage?: string | null;
  recordedAt?: string;
  lastSeenRunId?: string;
  createdAt?: string;
  updatedAt?: string;
};

type UserCalendarDetail = {
  target?: TargetProvider;
  targetUserId?: string;
  targetCalendarId?: string;
  destinationId: string;
  googleCalendarId: string;
  summary: string;
  description: string;
  timeZone: string;
  createdAt: string;
  updatedAt: string;
};

type UserDetailPayload = {
  user: Record<string, unknown>;
  events: DiagnosticEvent[];
  calendars: UserCalendarDetail[];
  runs: RunUserDiagnostic[];
  exclusions: UserEventExclusions;
  eventTypes: DiscoveredEventType[];
  globalPolicy: SyncPolicy;
};

const EVENT_CATEGORY_COPY: Record<EventCategory, [string, string]> = {
  timetable: ["Timetable lessons", "Classes and lessons identified by Schoolbox timetable metadata."],
  resource_booking: ["Resource bookings", "Rooms, equipment and other resource-linked bookings."],
  school_event: ["School events", "Items explicitly labelled as school-wide events."],
  individual_event: ["Individual events", "Personal or individual calendar items."],
  other: ["Other and custom", "Unclassified or installation-specific sources."],
};

const SCOPES = [
  "https://www.googleapis.com/auth/calendar.events.owned",
  "https://www.googleapis.com/auth/calendar.app.created",
  "https://www.googleapis.com/auth/admin.directory.user.readonly",
];

const SECONDARY_CALENDAR_SCOPE = "https://www.googleapis.com/auth/calendar.app.created";

const CALENDAR_COLOURS = [
  ["1", "Lavender"], ["2", "Sage"], ["3", "Grape"], ["4", "Flamingo"],
  ["5", "Banana"], ["6", "Tangerine"], ["7", "Peacock"], ["8", "Graphite"],
  ["9", "Blueberry"], ["10", "Basil"], ["11", "Tomato"],
] as const;

function useApplicationOrigin(): string {
  const [origin, setOrigin] = useState("");
  useEffect(() => {
    const timer = window.setTimeout(() => setOrigin(window.location.origin), 0);
    return () => window.clearTimeout(timer);
  }, []);
  return origin;
}

function normaliseCounts(value: unknown): TargetCounts {
  const row = value && typeof value === "object" ? value as Record<string, unknown> : {};
  return {
    users: Number(row.users ?? 0),
    enabled: Number(row.enabled ?? row.users ?? 0),
    disabled: Number(row.disabled ?? 0),
    healthy: Number(row.healthy ?? 0),
    errors: Number(row.errors ?? 0),
    unmatched: Number(row.unmatched ?? 0),
    events: Number(row.events ?? 0),
  };
}

function normalisePeople(value: unknown, fallbackTarget: TargetProvider = "google"): Person[] | null {
  if (!Array.isArray(value)) return null;
  return value.map((item, index) => {
    const row = item as Record<string, unknown>;
    const statusValue = String(row.status ?? "Syncing").toLowerCase();
    const status: Person["status"] = statusValue === "synced" ? "Synced" : statusValue === "pending" ? "Pending" : statusValue === "unmatched" ? "Unmatched" : statusValue === "error" || statusValue === "failed" ? "Error" : "Syncing";
    return {
      target: row.target === "microsoft" ? "microsoft" : fallbackTarget,
      id: String(row.id ?? row.targetUserId ?? row.googleUserId ?? `user-${index}`),
      name: String(row.name ?? row.displayName ?? "Unknown user"),
      schoolboxEmail: String(row.schoolboxEmail ?? row.sourceEmail ?? row.email ?? "—"),
      targetEmail: String(row.targetEmail ?? row.googleEmail ?? row.email ?? "—"),
      googleEmail: String(row.googleEmail ?? row.targetEmail ?? row.email ?? "—"),
      role: String(row.role ?? "User"),
      status,
      hasCustomExclusions: Boolean(row.hasCustomExclusions),
      syncEnabled: status !== "Unmatched" && (row.syncEnabled === undefined ? true : Boolean(row.syncEnabled)),
      eventCount: Math.max(0, Number(row.eventCount ?? 0)),
      calendarCount: Math.max(0, Number(row.calendarCount ?? 0)),
      lastSync: String(row.lastSync ?? row.lastSyncAt ?? row.last_synced_at ?? "Not yet"),
    };
  });
}

function normaliseRuns(value: unknown): Run[] | null {
  if (!Array.isArray(value)) return null;
  return value.map((item, index) => {
    const row = item as Record<string, unknown>;
    const rawStatus = String(row.status ?? "Succeeded").toLowerCase();
    const status: Run["status"] = rawStatus === "running" ? "Running" : rawStatus === "failed" ? "Failed" : rawStatus === "warning" || rawStatus === "completed_with_errors" ? "Warning" : "Succeeded";
    const created = Number(row.eventsCreated ?? 0);
    const updated = Number(row.eventsUpdated ?? 0);
    const deleted = Number(row.eventsDeleted ?? 0);
    const startedAt = String(row.started ?? row.startedAt ?? row.created_at ?? "Recently");
    const completedAt = String(row.completedAt ?? "");
    const durationMs = completedAt && !Number.isNaN(Date.parse(startedAt)) && !Number.isNaN(Date.parse(completedAt)) ? Date.parse(completedAt) - Date.parse(startedAt) : 0;
    const usersDiscovered = Number(row.usersDiscovered ?? row.users ?? 0);
    const targets = Array.isArray(row.targets) ? row.targets.flatMap((value) => {
      const targetRow = value as Record<string, unknown>;
      if (targetRow.target !== "google" && targetRow.target !== "microsoft") return [];
      return [{
        runId: String(targetRow.runId ?? row.id ?? ""),
        target: targetRow.target,
        status: String(targetRow.status ?? "pending"),
        phase: String(targetRow.phase ?? "pending"),
        phaseDetail: targetRow.phaseDetail ? String(targetRow.phaseDetail) : null,
        startedAt: String(targetRow.startedAt ?? startedAt),
        completedAt: targetRow.completedAt ? String(targetRow.completedAt) : null,
        usersDiscovered: Number(targetRow.usersDiscovered ?? 0),
        usersMatched: Number(targetRow.usersMatched ?? 0),
        usersSelected: Number(targetRow.usersSelected ?? 0),
        usersSynced: Number(targetRow.usersSynced ?? 0),
        eventsCreated: Number(targetRow.eventsCreated ?? 0),
        eventsUpdated: Number(targetRow.eventsUpdated ?? 0),
        eventsDeleted: Number(targetRow.eventsDeleted ?? 0),
        eventsUnchanged: Number(targetRow.eventsUnchanged ?? 0),
        errors: Number(targetRow.errors ?? 0),
        message: targetRow.message ? String(targetRow.message) : null,
      } satisfies RunTargetDiagnostic];
    }) : undefined;
    return {
      id: String(row.id ?? `RUN-${index + 1}`),
      started: !Number.isNaN(Date.parse(startedAt)) ? new Date(startedAt).toLocaleString("en-AU", { day: "numeric", month: "short", hour: "numeric", minute: "2-digit" }) : startedAt,
      trigger: String(row.trigger ?? "Scheduled"),
      status,
      users: Number(row.usersSynced ?? row.usersProcessed ?? row.users ?? usersDiscovered),
      usersDiscovered,
      usersMatched: Number(row.usersMatched ?? usersDiscovered),
      changes: Number(row.changes ?? row.eventsChanged ?? created + updated + deleted),
      duration: String(row.duration ?? (durationMs > 0 ? `${Math.floor(durationMs / 60000)}m ${Math.round(durationMs % 60000 / 1000)}s` : "—")),
      note: String(row.note ?? row.message ?? row.phaseDetail ?? "Run details are not available."),
      phase: String(row.phase ?? (status === "Running" ? "starting" : status === "Failed" ? "failed" : "completed")),
      phaseDetail: String(row.phaseDetail ?? row.message ?? "No phase detail is available."),
      progressAt: row.progressAt ? String(row.progressAt) : null,
      created,
      updated,
      deleted,
      errors: Number(row.errors ?? 0),
      targets,
    };
  });
}

let activeCsrfToken = "";
const UNAUTHORIZED_EVENT = "relay:unauthorized";
const SESSION_WARNING_MS = 5 * 60 * 1000;

class ApiError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
  }
}

async function fetchJson(url: string, init?: RequestInit) {
  const method = (init?.method ?? "GET").toUpperCase();
  const headers: Record<string, string> = { "Content-Type": "application/json" };
  if (!["GET", "HEAD", "OPTIONS"].includes(method) && activeCsrfToken) headers["X-CSRF-Token"] = activeCsrfToken;
  const response = await fetch(url, { ...init, credentials: "same-origin", headers: { ...headers, ...(init?.headers ?? {}) } });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 401 && url !== "/api/auth/login" && typeof window !== "undefined") {
      window.dispatchEvent(new Event(UNAUTHORIZED_EVENT));
    }
    throw new ApiError((data as { error?: string }).error || `Request failed (${response.status})`, response.status);
  }
  return data as Record<string, unknown>;
}

export default function Home() {
  const [auth, setAuth] = useState<AuthSession | null | undefined>(undefined);
  const [readiness, setReadiness] = useState<AuthReadiness>({ localAdministrator: false, googleSignInConfigured: false });
  const [authUnavailable, setAuthUnavailable] = useState(false);
  const [view, setView] = useState<View>("dashboard");
  const [mobileNav, setMobileNav] = useState(false);
  const [loading, setLoading] = useState(true);
  const [apiOnline, setApiOnline] = useState(false);
  const [configured, setConfigured] = useState(false);
  const [notice, setNotice] = useState<Notice>(null);
  const [peopleByTarget, setPeopleByTarget] = useState<Record<TargetProvider, Person[]>>({ google: [], microsoft: [] });
  const [runs, setRuns] = useState<Run[]>([]);
  const [selectedRun, setSelectedRun] = useState<Run | null>(null);
  const [syncRunning, setSyncRunning] = useState(false);
  const [manualTarget, setManualTarget] = useState<"all" | TargetProvider>("all");
  const [setupTrack, setSetupTrack] = useState<SetupTrack>("hub");
  const [setupResume, setSetupResume] = useState<SetupResume | undefined>(undefined);
  const [settingsInitialSection, setSettingsInitialSection] = useState<string | undefined>(undefined);
  const [peopleInitialTarget, setPeopleInitialTarget] = useState<TargetProvider | undefined>(undefined);
  const [lastSync, setLastSync] = useState("Never");
  const [health, setHealth] = useState("Setup required");
  const [counts, setCounts] = useState<TargetCounts | null>(null);
  const [targetCounts, setTargetCounts] = useState<Record<TargetProvider, TargetCounts>>({ google: { ...EMPTY_TARGET_COUNTS }, microsoft: { ...EMPTY_TARGET_COUNTS } });
  const [resourceErrors, setResourceErrors] = useState({ googlePeople: false, microsoftPeople: false, runs: false });
  const [loginMessage, setLoginMessage] = useState("");
  const [sessionClock, setSessionClock] = useState(() => Date.now());
  const [extendingSession, setExtendingSession] = useState(false);
  const [config, setConfig] = useState<Config>({
    schoolboxUrl: "",
    schoolboxJwt: "",
    serviceAccountJson: "",
    adminEmail: "",
    interval: "360",
    pastDays: "30",
    futureDays: "180",
    syncNewUsersByDefault: false,
    googleEnabled: true,
    syncNewGoogleUsersByDefault: false,
    googleCustomer: "my_customer",
    microsoftEnabled: false,
    microsoftTenantId: "",
    microsoftClientId: "",
    microsoftClientSecret: "",
    microsoftTestUserEmail: "",
    microsoftConsentGrantedAt: "",
    hasMicrosoftClientSecret: false,
    syncNewMicrosoftUsersByDefault: false,
    timezone: "Australia/Sydney",
    concurrency: "3",
    discoveryTimeoutSeconds: "120",
    userSyncTimeoutSeconds: "180",
    runTimeoutMinutes: "30",
    enabled: false,
    setupCompleted: false,
    schoolboxSetupCompleted: false,
    googleSetupCompleted: false,
    microsoftSetupCompleted: false,
    hasSchoolboxToken: false,
    hasGoogleServiceAccount: false,
    serviceAccountEmail: "",
    serviceAccountClientId: "",
    syncPolicy: normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY),
    microsoftSyncPolicy: normalizeSyncPolicy({}, DEFAULT_SYNC_POLICY),
  });

  const people = useMemo(() => [...peopleByTarget.google, ...peopleByTarget.microsoft], [peopleByTarget]);

  useEffect(() => {
    if ((manualTarget === "google" && config.googleEnabled) || (manualTarget === "microsoft" && config.microsoftEnabled) || manualTarget === "all") return;
    const timer = window.setTimeout(() => setManualTarget("all"), 0);
    return () => window.clearTimeout(timer);
  }, [manualTarget, config.googleEnabled, config.microsoftEnabled]);

  const canOperate = Boolean(auth?.permissions.includes("operate"));
  const canConfigure = Boolean(auth?.permissions.includes("configure"));
  const canManageAccess = Boolean(auth?.permissions.includes("manage_access"));

  const expireSession = useCallback(() => {
    activeCsrfToken = "";
    setAuth(null);
    setView("dashboard");
    setLoginMessage("Your Relay session has expired. Sign in again to continue.");
    window.history.replaceState({}, "", "/");
  }, []);

  const acceptSession = useCallback((session: AuthSession) => {
    activeCsrfToken = session.csrfToken;
    setAuth(session);
    setView("dashboard");
    setLoginMessage("");
    setSessionClock(Date.now());
  }, []);

  useEffect(() => {
    let cancelled = false;
    const callbackParams = new URLSearchParams(window.location.search);
    const callbackVerified = callbackParams.get("microsoftConsent") === "verified";
    const callbackError = callbackParams.get("microsoftConsentError");
    if (callbackVerified || callbackError) {
      window.sessionStorage.setItem("relay:microsoft-consent-result", JSON.stringify({ verified: callbackVerified, error: callbackError }));
    }
    void fetchJson("/api/auth/session")
      .then((payload) => {
        if (cancelled) return;
        setReadiness((payload.readiness as AuthReadiness | undefined) ?? { localAdministrator: false, googleSignInConfigured: false });
        setAuthUnavailable(false);
        if (payload.authenticated && payload.session) acceptSession(payload.session as AuthSession);
        else {
          activeCsrfToken = "";
          setAuth(null);
        }
      })
      .catch(() => { if (!cancelled) { setAuthUnavailable(true); setAuth(null); } });
    return () => { cancelled = true; };
  }, [acceptSession]);

  useEffect(() => {
    const handleUnauthorized = () => expireSession();
    window.addEventListener(UNAUTHORIZED_EVENT, handleUnauthorized);
    return () => window.removeEventListener(UNAUTHORIZED_EVENT, handleUnauthorized);
  }, [expireSession]);

  const authenticatedUserId = auth?.userId;
  useEffect(() => {
    if (!authenticatedUserId) return;
    let cancelled = false;
    const refreshSessionState = async () => {
      try {
        const payload = await fetchJson("/api/auth/session");
        if (cancelled) return;
        if (!payload.authenticated || !payload.session) {
          expireSession();
          return;
        }
        const session = payload.session as AuthSession;
        activeCsrfToken = session.csrfToken;
        setAuth(session);
      } catch (error) {
        if (!(error instanceof ApiError) || error.status !== 401) return;
      }
    };
    const timer = window.setInterval(() => void refreshSessionState(), 30_000);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [authenticatedUserId, expireSession]);

  useEffect(() => {
    if (!authenticatedUserId) return;
    const timer = window.setInterval(() => setSessionClock(Date.now()), 1_000);
    return () => window.clearInterval(timer);
  }, [authenticatedUserId]);

  const loadData = useCallback(async () => {
    setLoading(true);
    setApiOnline(false);
    const results = await Promise.allSettled([
      fetchJson("/api/status"),
      fetchJson("/api/config"),
      fetchJson("/api/diagnostics"),
      fetchJson("/api/users?target=google"),
      fetchJson("/api/users?target=microsoft"),
      fetchJson("/api/runs?limit=30"),
    ]);
    const statusResult = results[0];
    const configResult = results[1];
    const diagnosticResult = results[2];
    if (statusResult.status === "fulfilled") {
      setApiOnline(true);
      const payload = statusResult.value;
      const lastRun = payload.lastRun as Record<string, unknown> | undefined;
      const statusConfig = payload.schedule as Record<string, unknown> | undefined;
      setConfigured(Boolean(payload.configured));
      if (payload.counts) setCounts(normaliseCounts(payload.counts));
      const incomingTargetCounts = payload.targetCounts as Partial<Record<TargetProvider, unknown>> | undefined;
      if (incomingTargetCounts) setTargetCounts({
        google: normaliseCounts(incomingTargetCounts.google),
        microsoft: normaliseCounts(incomingTargetCounts.microsoft),
      });
      setHealth(String(payload.health ?? payload.status ?? (!payload.configured ? "Setup required" : lastRun?.status === "failed" ? "Failed" : lastRun?.status === "completed_with_errors" ? "Warning" : "Healthy")));
      const syncDate = payload.lastSync ?? payload.last_sync_at ?? lastRun?.completedAt ?? lastRun?.startedAt;
      setLastSync(syncDate ? new Date(String(syncDate)).toLocaleString("en-AU") : "Never");
      if (statusConfig) setConfig(current => ({
        ...current,
        interval: String(statusConfig.syncIntervalMinutes ?? current.interval),
        pastDays: String(statusConfig.pastDays ?? current.pastDays),
        futureDays: String(statusConfig.futureDays ?? current.futureDays),
        discoveryTimeoutSeconds: String(statusConfig.discoveryTimeoutSeconds ?? current.discoveryTimeoutSeconds),
        userSyncTimeoutSeconds: String(statusConfig.userSyncTimeoutSeconds ?? current.userSyncTimeoutSeconds),
        runTimeoutMinutes: String(statusConfig.runTimeoutMinutes ?? current.runTimeoutMinutes),
        googleEnabled: Boolean(statusConfig.googleEnabled ?? current.googleEnabled),
        syncNewGoogleUsersByDefault: Boolean(statusConfig.syncNewGoogleUsersByDefault ?? statusConfig.syncNewUsersByDefault ?? current.syncNewGoogleUsersByDefault),
        syncNewUsersByDefault: Boolean(statusConfig.syncNewGoogleUsersByDefault ?? statusConfig.syncNewUsersByDefault ?? current.syncNewUsersByDefault),
        microsoftEnabled: Boolean(statusConfig.microsoftEnabled ?? current.microsoftEnabled),
        syncNewMicrosoftUsersByDefault: Boolean(statusConfig.syncNewMicrosoftUsersByDefault ?? current.syncNewMicrosoftUsersByDefault),
        microsoftConsentGrantedAt: String(statusConfig.microsoftConsentGrantedAt ?? current.microsoftConsentGrantedAt),
        timezone: String(statusConfig.timezone ?? current.timezone),
        enabled: Boolean(statusConfig.enabled ?? current.enabled),
        setupCompleted: Boolean(statusConfig.setupCompleted ?? current.setupCompleted),
        schoolboxSetupCompleted: Boolean(statusConfig.schoolboxSetupCompleted ?? current.schoolboxSetupCompleted),
        googleSetupCompleted: Boolean(statusConfig.googleSetupCompleted ?? current.googleSetupCompleted),
        microsoftSetupCompleted: Boolean(statusConfig.microsoftSetupCompleted ?? current.microsoftSetupCompleted),
        syncPolicy: normalizeSyncPolicy(statusConfig.syncPolicy, current.syncPolicy),
        microsoftSyncPolicy: normalizeSyncPolicy(statusConfig.microsoftSyncPolicy, current.microsoftSyncPolicy),
      }));
      const fetchedRuns = normaliseRuns(payload.runs ?? payload.history ?? (payload.lastRun ? [payload.lastRun] : null));
      if (fetchedRuns) setRuns(fetchedRuns);
    } else if (statusResult.reason instanceof ApiError && statusResult.reason.status === 401) {
      activeCsrfToken = "";
      setAuth(null);
    }
    if (configResult.status === "fulfilled") {
      const payload = configResult.value;
      const incoming = (payload.config ?? payload) as Record<string, unknown>;
      setConfig(current => ({
        ...current,
        schoolboxUrl: String(incoming.schoolboxUrl ?? incoming.schoolboxBaseUrl ?? incoming.schoolbox_url ?? current.schoolboxUrl),
        adminEmail: String(incoming.adminEmail ?? incoming.googleAdminEmail ?? incoming.delegatedAdminEmail ?? incoming.admin_email ?? current.adminEmail),
        interval: String(incoming.interval ?? incoming.syncIntervalMinutes ?? current.interval),
        pastDays: String(incoming.pastDays ?? incoming.past_days ?? current.pastDays),
        futureDays: String(incoming.futureDays ?? incoming.future_days ?? current.futureDays),
        googleEnabled: Boolean(incoming.googleEnabled ?? current.googleEnabled),
        syncNewGoogleUsersByDefault: Boolean(incoming.syncNewGoogleUsersByDefault ?? incoming.syncNewUsersByDefault ?? current.syncNewGoogleUsersByDefault),
        syncNewUsersByDefault: Boolean(incoming.syncNewGoogleUsersByDefault ?? incoming.syncNewUsersByDefault ?? current.syncNewUsersByDefault),
        googleCustomer: String(incoming.googleCustomer ?? current.googleCustomer),
        microsoftEnabled: Boolean(incoming.microsoftEnabled ?? current.microsoftEnabled),
        microsoftTenantId: String(incoming.microsoftTenantId ?? current.microsoftTenantId),
        microsoftClientId: String(incoming.microsoftClientId ?? current.microsoftClientId),
        microsoftTestUserEmail: String(incoming.microsoftTestUserEmail ?? current.microsoftTestUserEmail),
        microsoftConsentGrantedAt: String(incoming.microsoftConsentGrantedAt ?? ""),
        hasMicrosoftClientSecret: Boolean(incoming.hasMicrosoftClientSecret ?? current.hasMicrosoftClientSecret),
        syncNewMicrosoftUsersByDefault: Boolean(incoming.syncNewMicrosoftUsersByDefault ?? current.syncNewMicrosoftUsersByDefault),
        timezone: String(incoming.timezone ?? current.timezone),
        concurrency: String(incoming.concurrency ?? current.concurrency),
        discoveryTimeoutSeconds: String(incoming.discoveryTimeoutSeconds ?? current.discoveryTimeoutSeconds),
        userSyncTimeoutSeconds: String(incoming.userSyncTimeoutSeconds ?? current.userSyncTimeoutSeconds),
        runTimeoutMinutes: String(incoming.runTimeoutMinutes ?? current.runTimeoutMinutes),
        enabled: Boolean(incoming.enabled ?? current.enabled),
        setupCompleted: Boolean(incoming.setupCompleted ?? current.setupCompleted),
        schoolboxSetupCompleted: Boolean(incoming.schoolboxSetupCompleted ?? current.schoolboxSetupCompleted),
        googleSetupCompleted: Boolean(incoming.googleSetupCompleted ?? current.googleSetupCompleted),
        microsoftSetupCompleted: Boolean(incoming.microsoftSetupCompleted ?? current.microsoftSetupCompleted),
        hasSchoolboxToken: Boolean(incoming.hasSchoolboxToken ?? current.hasSchoolboxToken),
        hasGoogleServiceAccount: Boolean(incoming.hasGoogleServiceAccount ?? current.hasGoogleServiceAccount),
        serviceAccountEmail: String(incoming.serviceAccountEmail ?? current.serviceAccountEmail),
        serviceAccountClientId: String(incoming.serviceAccountClientId ?? current.serviceAccountClientId),
        syncPolicy: normalizeSyncPolicy(incoming.syncPolicy, current.syncPolicy),
        microsoftSyncPolicy: normalizeSyncPolicy(incoming.microsoftSyncPolicy, current.microsoftSyncPolicy),
      }));
    }
    if (diagnosticResult.status === "fulfilled") {
      const payload = diagnosticResult.value;
      const fetchedRuns = normaliseRuns(payload.runs ?? payload.history);
      if (fetchedRuns?.length) setRuns(fetchedRuns);
    }
    const googleUsersResult = results[3];
    const microsoftUsersResult = results[4];
    setResourceErrors({ googlePeople: googleUsersResult.status === "rejected", microsoftPeople: microsoftUsersResult.status === "rejected", runs: results[5].status === "rejected" });
    if (googleUsersResult.status === "fulfilled") {
      const fetchedPeople = normalisePeople(googleUsersResult.value.users, "google");
      if (fetchedPeople) setPeopleByTarget(current => ({ ...current, google: fetchedPeople }));
    }
    if (microsoftUsersResult.status === "fulfilled") {
      const fetchedPeople = normalisePeople(microsoftUsersResult.value.users, "microsoft");
      if (fetchedPeople) setPeopleByTarget(current => ({ ...current, microsoft: fetchedPeople }));
    }
    const runsResult = results[5];
    if (runsResult.status === "fulfilled") {
      const fetchedRuns = normaliseRuns(runsResult.value.runs);
      if (fetchedRuns) setRuns(fetchedRuns);
    }
    setLoading(false);
  }, []);

  // Session polling replaces the auth object every 30 seconds. Key this
  // initial load to the stable user ID so routine refreshes preserve drafts.
  useEffect(() => {
    if (!authenticatedUserId) return;
    const timer = window.setTimeout(() => void loadData(), 0);
    return () => window.clearTimeout(timer);
  }, [authenticatedUserId, loadData]);
  useEffect(() => {
    if (!authenticatedUserId) return;
    const params = new URLSearchParams(window.location.search);
    let storedResult: { verified?: boolean; error?: string | null } = {};
    try { storedResult = JSON.parse(window.sessionStorage.getItem("relay:microsoft-consent-result") || "{}"); } catch { /* Ignore malformed browser state. */ }
    const verified = params.get("microsoftConsent") === "verified" || Boolean(storedResult.verified);
    const consentError = params.get("microsoftConsentError") || storedResult.error || null;
    if (!verified && !consentError) return;
    params.delete("microsoftConsent");
    params.delete("microsoftConsentError");
    const query = params.toString();
    window.history.replaceState({}, "", `${window.location.pathname}${query ? `?${query}` : ""}`);
    const timer = window.setTimeout(() => void (async () => {
      if (verified) setNotice({ kind: "success", message: "Microsoft tenant admin consent and Microsoft Graph directory/calendar access were verified successfully." });
      else setNotice({ kind: "error", message: `Microsoft admin consent did not complete${consentError ? `: ${consentError}` : "."}` });
      window.sessionStorage.removeItem("relay:microsoft-consent-result");
      const savedProgressText = window.sessionStorage.getItem("relay:microsoft-consent-progress");
      if (!savedProgressText) {
        window.sessionStorage.removeItem("relay:microsoft-consent-progress");
        setSettingsInitialSection("Connections");
        setView("settings");
      } else {
        let progress: SetupResume = { schoolboxTested: false, googleTested: false, microsoftTested: false };
        try {
          const savedProgress = JSON.parse(savedProgressText || "{}");
          progress = {
            schoolboxTested: Boolean(savedProgress.schoolboxTested),
            googleTested: Boolean(savedProgress.googleTested),
            microsoftTested: verified,
            microsoftEnabled: Boolean(savedProgress.microsoftEnabled),
            syncNewMicrosoftUsersByDefault: Boolean(savedProgress.syncNewMicrosoftUsersByDefault),
          };
        } catch { progress.microsoftTested = verified; }
        await loadData();
        window.sessionStorage.removeItem("relay:microsoft-consent-progress");
        setSetupResume(progress);
        setSetupTrack("microsoft");
        setView("setup");
      }
      if (!savedProgressText) await loadData();
    })(), 0);
    return () => window.clearTimeout(timer);
  }, [authenticatedUserId, loadData]);
  useEffect(() => {
    if (!notice) return;
    const timer = window.setTimeout(() => setNotice(null), 5000);
    return () => window.clearTimeout(timer);
  }, [notice]);

  const sessionDeadlines = auth
    ? [Date.parse(auth.expiresAt), Date.parse(auth.idleExpiresAt)].filter(Number.isFinite)
    : [];
  const sessionDeadline = sessionDeadlines.length ? Math.min(...sessionDeadlines) : Number.POSITIVE_INFINITY;
  const sessionRemainingMs = sessionDeadline - sessionClock;
  const showSessionWarning = Boolean(auth && Number.isFinite(sessionDeadline) && sessionRemainingMs <= SESSION_WARNING_MS);

  const changeView = (next: View) => {
    setView(next);
    setMobileNav(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const runNow = async (targets?: TargetProvider[]) => {
    setSyncRunning(true);
    const selectedLabels = targets?.map(target => TARGET_LABELS[target]).join(" and ");
    setNotice({ kind: "info", message: `Discovering users and syncing ${selectedLabels ?? "all enabled targets"}…` });
    try {
      const data = await fetchJson("/api/sync/run", { method: "POST", body: JSON.stringify({ trigger: "manual", ...(targets ? { targets } : {}) }) });
      setLastSync("Just now");
      setNotice({ kind: "success", message: String(data.message ?? "Sync started. Progress will appear in Runs.") });
      await loadData();
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) setAuth(null);
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "The sync could not be started." });
    } finally {
      setSyncRunning(false);
    }
  };

  const saveConfig = async (message = "Settings saved", forceEnabled?: boolean) => {
    try {
      const saved = await fetchJson("/api/config", {
        method: "PUT",
        body: JSON.stringify({
          schoolboxBaseUrl: config.schoolboxUrl,
          ...(config.schoolboxJwt ? { schoolboxToken: config.schoolboxJwt } : {}),
          ...(config.serviceAccountJson ? { googleServiceAccountJson: config.serviceAccountJson } : {}),
          googleAdminEmail: config.adminEmail,
          googleEnabled: config.googleEnabled,
          syncIntervalMinutes: Number(config.interval),
          pastDays: Number(config.pastDays),
          futureDays: Number(config.futureDays),
          syncNewGoogleUsersByDefault: config.syncNewGoogleUsersByDefault,
          syncNewUsersByDefault: config.syncNewGoogleUsersByDefault,
          syncPolicy: config.syncPolicy,
          googleCustomer: config.googleCustomer,
          microsoftEnabled: config.microsoftEnabled,
          microsoftTenantId: config.microsoftTenantId,
          microsoftClientId: config.microsoftClientId,
          ...(config.microsoftClientSecret ? { microsoftClientSecret: config.microsoftClientSecret } : {}),
          microsoftTestUserEmail: config.microsoftTestUserEmail,
          syncNewMicrosoftUsersByDefault: config.syncNewMicrosoftUsersByDefault,
          microsoftSyncPolicy: config.microsoftSyncPolicy,
          concurrency: Number(config.concurrency),
          discoveryTimeoutSeconds: Number(config.discoveryTimeoutSeconds),
          userSyncTimeoutSeconds: Number(config.userSyncTimeoutSeconds),
          runTimeoutMinutes: Number(config.runTimeoutMinutes),
          enabled: forceEnabled ?? config.enabled,
          timezone: config.timezone,
        }),
      });
      setApiOnline(true);
      setConfigured(Boolean(saved.setupCompleted));
      setConfig(current => ({
        ...current,
        schoolboxJwt: "",
        serviceAccountJson: "",
        microsoftClientSecret: "",
        hasSchoolboxToken: Boolean(saved.hasSchoolboxToken ?? (current.hasSchoolboxToken || Boolean(current.schoolboxJwt))),
        hasGoogleServiceAccount: Boolean(saved.hasGoogleServiceAccount ?? (current.hasGoogleServiceAccount || Boolean(current.serviceAccountJson))),
        serviceAccountEmail: String(saved.serviceAccountEmail ?? current.serviceAccountEmail),
        serviceAccountClientId: String(saved.serviceAccountClientId ?? (() => {
          try { return JSON.parse(current.serviceAccountJson || "{}").client_id ?? current.serviceAccountClientId; }
          catch { return current.serviceAccountClientId; }
        })()),
        googleEnabled: Boolean(saved.googleEnabled ?? current.googleEnabled),
        syncNewGoogleUsersByDefault: Boolean(saved.syncNewGoogleUsersByDefault ?? saved.syncNewUsersByDefault ?? current.syncNewGoogleUsersByDefault),
        syncNewUsersByDefault: Boolean(saved.syncNewGoogleUsersByDefault ?? saved.syncNewUsersByDefault ?? current.syncNewUsersByDefault),
        googleCustomer: String(saved.googleCustomer ?? current.googleCustomer),
        microsoftEnabled: Boolean(saved.microsoftEnabled ?? current.microsoftEnabled),
        microsoftTenantId: String(saved.microsoftTenantId ?? current.microsoftTenantId),
        microsoftClientId: String(saved.microsoftClientId ?? current.microsoftClientId),
        microsoftTestUserEmail: String(saved.microsoftTestUserEmail ?? current.microsoftTestUserEmail),
        microsoftConsentGrantedAt: String(saved.microsoftConsentGrantedAt ?? ""),
        hasMicrosoftClientSecret: Boolean(saved.hasMicrosoftClientSecret ?? (current.hasMicrosoftClientSecret || Boolean(current.microsoftClientSecret))),
        syncNewMicrosoftUsersByDefault: Boolean(saved.syncNewMicrosoftUsersByDefault ?? current.syncNewMicrosoftUsersByDefault),
        timezone: String(saved.timezone ?? current.timezone),
        concurrency: String(saved.concurrency ?? current.concurrency),
        discoveryTimeoutSeconds: String(saved.discoveryTimeoutSeconds ?? current.discoveryTimeoutSeconds),
        userSyncTimeoutSeconds: String(saved.userSyncTimeoutSeconds ?? current.userSyncTimeoutSeconds),
        runTimeoutMinutes: String(saved.runTimeoutMinutes ?? current.runTimeoutMinutes),
        enabled: Boolean(saved.enabled ?? current.enabled),
        setupCompleted: Boolean(saved.setupCompleted ?? current.setupCompleted),
        schoolboxSetupCompleted: Boolean(saved.schoolboxSetupCompleted ?? current.schoolboxSetupCompleted),
        googleSetupCompleted: Boolean(saved.googleSetupCompleted ?? current.googleSetupCompleted),
        microsoftSetupCompleted: Boolean(saved.microsoftSetupCompleted ?? current.microsoftSetupCompleted),
        syncPolicy: normalizeSyncPolicy(saved.syncPolicy, current.syncPolicy),
        microsoftSyncPolicy: normalizeSyncPolicy(saved.microsoftSyncPolicy, current.microsoftSyncPolicy),
      }));
      setNotice({ kind: "success", message });
      return true;
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) setAuth(null);
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Settings could not be saved." });
      return false;
    }
  };

  const extendCurrentSession = async () => {
    setExtendingSession(true);
    try {
      const payload = await fetchJson("/api/auth/session", { method: "POST", body: "{}" });
      const session = payload.session as AuthSession;
      activeCsrfToken = session.csrfToken;
      setAuth(session);
      setSessionClock(Date.now());
      setNotice({ kind: "success", message: "Your Relay session has been extended." });
    } catch (error) {
      if (!(error instanceof ApiError) || error.status !== 401) {
        setNotice({ kind: "error", message: error instanceof Error ? error.message : "The session could not be extended." });
      }
    } finally {
      setExtendingSession(false);
    }
  };

  const signOut = async () => {
    try {
      await fetchJson("/api/auth/logout", { method: "POST", body: "{}" });
      activeCsrfToken = "";
      setAuth(null);
      setView("dashboard");
      setLoginMessage("");
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? `Sign out failed: ${error.message}` : "Sign out failed. Try again." });
    }
  };
  const handleSignedOut = useCallback(() => {
    activeCsrfToken = "";
    setAuth(null);
    setView("dashboard");
    setLoginMessage("Sign in again to continue.");
  }, []);

  if (auth === undefined) return <div className="auth-shell"><div className="auth-loading"><span className="brand-mark" aria-hidden="true"><i /><i /><i /></span><p>Starting Relay…</p></div></div>;
  if (!auth) return <LoginScreen readiness={readiness} unavailable={authUnavailable} message={loginMessage} onAuthenticated={acceptSession} />;

  const title = { dashboard: "Calendar operations", setup: "Connection setup", people: "People & sync coverage", runs: "Runs & troubleshooting", settings: "Sync settings", access: "IT access" }[view];
  const subtitle = {
    dashboard: configured ? "Monitor Schoolbox calendar delivery across every configured target." : "Complete setup before discovering users and starting calendar sync.",
    setup: "Configure Schoolbox, Google Workspace and Microsoft 365 as independent connection tracks.",
    people: "Review each target directory independently and choose whose calendars Relay maintains.",
    runs: "Inspect every sync and find the cause of exceptions.",
    settings: "Control schedule, calendar coverage and operational alerts.",
    access: "Configure Google sign-in and control who can administer Relay.",
  }[view];
  const initials = auth.displayName.split(/\s+/).map(part => part[0]).join("").slice(0, 2).toUpperCase() || "A";
  const roleLabel = auth.isOwner ? "Local administrator" : auth.role === "admin" ? "Administrator" : auth.role === "operator" ? "Operator" : "Viewer";

  return (
    <div className="app-shell">
      {showSessionWarning && <div className="session-dialog-backdrop"><section className="session-dialog" role="dialog" aria-modal="true" aria-labelledby="session-dialog-title"><span className="session-dialog-icon">⌛</span><div><p className="eyebrow">Session expiring</p><h2 id="session-dialog-title">Stay signed in?</h2><p>Your Relay session will expire in about {Math.max(1, Math.ceil(sessionRemainingMs / 60_000))} minute{Math.ceil(sessionRemainingMs / 60_000) === 1 ? "" : "s"}. Extend it now to keep working.</p></div><div className="session-dialog-actions"><button className="button ghost" onClick={() => void signOut()} disabled={extendingSession}>Sign out</button><button className="button primary" onClick={() => void extendCurrentSession()} disabled={extendingSession}>{extendingSession ? "Extending…" : "Stay signed in"}</button></div></section></div>}
      <aside className={`sidebar ${mobileNav ? "sidebar-open" : ""}`}>
        <button className="brand" onClick={() => changeView("dashboard")} aria-label="Relay home">
          <span className="brand-mark" aria-hidden="true"><i /><i /><i /></span>
          <span><b>Relay</b><small>Calendar operations</small></span>
        </button>
        <nav aria-label="Main navigation">
          <p className="nav-label">Workspace</p>
          <NavButton active={view === "dashboard"} icon="⌂" label="Overview" onClick={() => changeView("dashboard")} />
          {canConfigure && <NavButton active={view === "setup"} icon={configured ? "✓" : "↗"} label="Connections" onClick={() => { setSetupTrack("hub"); setSetupResume(undefined); changeView("setup"); }} />}
          <NavButton active={view === "people"} icon="◎" label="People" count={counts?.users ? String(counts.users) : undefined} onClick={() => changeView("people")} />
          <NavButton active={view === "runs"} icon="≡" label="Runs" onClick={() => changeView("runs")} />
          {canConfigure && <NavButton active={view === "settings"} icon="⚙" label="Settings" onClick={() => changeView("settings")} />}
          {canManageAccess && <NavButton active={view === "access"} icon="◇" label="IT access" onClick={() => changeView("access")} />}
        </nav>
        <div className="sidebar-foot">
          <div className="mini-status"><span className={apiOnline ? "pulse" : "pulse offline"} /><span><b>{configured ? "Relay is active" : "Setup required"}</b><small>{apiOnline ? "Services responding" : "Service unavailable"}</small></span></div>
          <div className="profile"><span className="avatar">{initials}</span><span><b>{auth.displayName}</b><small>{roleLabel}</small></span><button className="profile-logout" onClick={() => void signOut()}>Sign out</button></div>
        </div>
      </aside>

      <main className="main">
        <header className="mobile-header">
          <button className="brand" onClick={() => changeView("dashboard")} aria-label="Relay home"><span className="brand-mark" aria-hidden="true"><i /><i /><i /></span><b>Relay</b></button>
          <button className="menu-button" onClick={() => setMobileNav(!mobileNav)} aria-expanded={mobileNav} aria-label="Toggle navigation">{mobileNav ? "×" : "☰"}</button>
        </header>
        <div className="main-inner">
          <div className="page-heading">
            <div><p className="eyebrow">Schoolbox <span>→</span> Organisation calendars</p><h1>{title}</h1><p>{subtitle}</p></div>
            {view !== "setup" && view !== "access" && <div className="heading-actions"><button className="button ghost" onClick={() => void loadData()} disabled={loading}>{loading ? "Refreshing…" : "Refresh"}</button>{canOperate && <><select className="run-target-select" value={manualTarget} onChange={event => setManualTarget(event.target.value as "all" | TargetProvider)} aria-label="Calendar targets for manual sync" disabled={syncRunning || !configured}><option value="all">All enabled targets</option>{config.googleEnabled && <option value="google">Google Workspace only</option>}{config.microsoftEnabled && <option value="microsoft">Microsoft 365 only</option>}</select><button className="button primary" onClick={() => void runNow(manualTarget === "all" ? undefined : [manualTarget])} disabled={syncRunning || !configured}>{syncRunning ? "Starting…" : "Run sync now"}<span aria-hidden="true">→</span></button></>}</div>}
          </div>

          {notice && <div role="status" className={`notice ${notice.kind}`}><span>{notice.kind === "success" ? "✓" : notice.kind === "error" ? "!" : "i"}</span>{notice.message}<button onClick={() => setNotice(null)} aria-label="Dismiss notification">×</button></div>}
          {view === "dashboard" && <Dashboard people={people} peopleByTarget={peopleByTarget} targetCounts={targetCounts} runs={runs} counts={counts} lastSync={lastSync} health={health} apiOnline={apiOnline} configured={configured} config={config} runsError={resourceErrors.runs} onNavigate={changeView} onOpenPeople={target => { setPeopleInitialTarget(target); changeView("people"); }} onOpenConnections={() => { setSettingsInitialSection("Connections"); changeView("settings"); }} onSelectRun={(run) => { setSelectedRun(run); setView("runs"); }} />}
          {view === "setup" && canConfigure && <SetupWizard key={setupResume ? "microsoft-consent-resume" : "setup-default"} configured={configured} track={setupTrack} setTrack={setSetupTrack} resume={setupResume} config={config} setConfig={setConfig} setConfigured={setConfigured} setNotice={setNotice} changeView={changeView} />}
          {view === "people" && <PeoplePage initialTarget={peopleInitialTarget} peopleByTarget={peopleByTarget} setPeopleByTarget={setPeopleByTarget} targetCounts={targetCounts} loadErrors={{ google: resourceErrors.googlePeople, microsoft: resourceErrors.microsoftPeople }} configuredTargets={{ google: config.googleEnabled, microsoft: config.microsoftEnabled }} canConfigure={canConfigure} setNotice={setNotice} onOpenConnections={() => { setSettingsInitialSection("Connections"); changeView("settings"); }} />}
          {view === "runs" && <RunsPage runs={runs} selectedRun={selectedRun} setSelectedRun={setSelectedRun} runNow={runNow} syncRunning={syncRunning} canOperate={canOperate} loadError={resourceErrors.runs} />}
          {view === "settings" && canConfigure && <SettingsPage initialSection={settingsInitialSection} config={config} setConfig={setConfig} saveConfig={saveConfig} setNotice={setNotice} onOpenSetup={target => { setSetupResume(undefined); setSetupTrack(target); changeView("setup"); }} />}
          {view === "access" && canManageAccess && <AccessPage canChangeLocalPassword={auth.isOwner && auth.authType === "local"} setNotice={setNotice} onSignedOut={handleSignedOut} />}
        </div>
      </main>
    </div>
  );
}

function LoginScreen({ readiness, unavailable, message, onAuthenticated }: { readiness: AuthReadiness; unavailable: boolean; message: string; onAuthenticated: (session: AuthSession) => void }) {
  const [username, setUsername] = useState("administrator");
  const [password, setPassword] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [callbackError, setCallbackError] = useState("");
  useEffect(() => {
    const timer = window.setTimeout(() => setCallbackError(new URLSearchParams(window.location.search).get("authError") ?? ""), 0);
    return () => window.clearTimeout(timer);
  }, []);

  const submit = async (event: FormEvent) => {
    event.preventDefault();
    setBusy(true);
    setError("");
    try {
      const payload = await fetchJson("/api/auth/login", { method: "POST", body: JSON.stringify({ username, password }) });
      onAuthenticated(payload.session as AuthSession);
      window.history.replaceState({}, "", "/");
    } catch (problem) {
      setError(problem instanceof Error ? problem.message : "Sign-in failed");
    } finally {
      setBusy(false);
    }
  };

  return <div className="auth-shell">
    <div className="auth-card">
      <div className="auth-brand"><span className="brand-mark" aria-hidden="true"><i /><i /><i /></span><div><b>Relay</b><small>Calendar operations</small></div></div>
      <p className="eyebrow">Internal administration</p>
      <h1>Sign in to Relay</h1>
      <p className="auth-intro">Use the break-glass administrator account, or sign in with an approved Google Workspace IT account.</p>
      {message && <div className="auth-info" role="status"><span>i</span>{message}</div>}
      {unavailable && <div className="auth-error" role="alert"><span>!</span><div>Relay could not reach its authentication service. <button onClick={() => window.location.reload()}>Try again</button></div></div>}
      {(error || callbackError) && <div className="auth-error" role="alert"><span>!</span>{error || callbackError}</div>}
      {!unavailable && readiness.localAdministrator ? <form onSubmit={submit} className="auth-form">
        <Field label="Username"><input autoComplete="username" value={username} onChange={event => setUsername(event.target.value)} required /></Field>
        <Field label="Password"><input type="password" autoComplete="current-password" value={password} onChange={event => setPassword(event.target.value)} required /></Field>
        <button className="button primary full" type="submit" disabled={busy}>{busy ? "Signing in…" : "Sign in as local administrator"}<span>→</span></button>
      </form> : !unavailable && <div className="auth-setup"><b>Local administrator setup is required</b><p>On the server, run <code>npm run auth:bootstrap</code> before opening Relay to the IT network.</p></div>}
      {!unavailable && readiness.googleSignInConfigured && <div className="oauth-choice"><span>or</span><a className="button google-signin" href="/api/auth/google/start"><b>G</b> Continue with Google Workspace</a><small>Only accounts pre-approved by a Relay administrator can enter.</small></div>}
      <div className="auth-foot"><span>🔒</span><p>Self-hosted on your internal server. Sessions expire after 30 minutes of inactivity.</p></div>
    </div>
  </div>;
}

function NavButton({ active, icon, label, count, onClick }: { active: boolean; icon: string; label: string; count?: string; onClick: () => void }) {
  return <button className={`nav-button ${active ? "active" : ""}`} onClick={onClick}><span className="nav-icon" aria-hidden="true">{icon}</span><span>{label}</span>{count && <small>{count}</small>}</button>;
}

function Dashboard({ people, peopleByTarget, targetCounts, runs, counts, lastSync, health, apiOnline, configured, config, runsError, onNavigate, onOpenPeople, onOpenConnections, onSelectRun }: { people: Person[]; peopleByTarget: Record<TargetProvider, Person[]>; targetCounts: Record<TargetProvider, TargetCounts>; runs: Run[]; counts: TargetCounts | null; lastSync: string; health: string; apiOnline: boolean; configured: boolean; config: Config; runsError: boolean; onNavigate: (view: View) => void; onOpenPeople: (target: TargetProvider) => void; onOpenConnections: () => void; onSelectRun: (run: Run) => void }) {
  const enabledPeople = people.filter(person => person.syncEnabled);
  const totalUsers = people.length ? enabledPeople.length : counts?.enabled ?? 0;
  const healthyUsers = people.length ? enabledPeople.filter(person => person.status === "Synced" || person.status === "Syncing").length : counts?.healthy ?? 0;
  const unmatchedUsers = people.length ? people.filter(person => person.status === "Unmatched").length : counts?.unmatched ?? 0;
  const latestRun = runs[0];
  const activity = [...runs.slice(0, 7)].reverse();
  const activityMaximum = Math.max(1, ...activity.flatMap(run => [run.created ?? 0, run.updated ?? 0]));
  const attention = health.toLowerCase().includes("fail") || health.toLowerCase().includes("warning");
  const healthTitle = !apiOnline ? "Service unavailable" : !configured ? "Setup is not complete" : attention ? "Attention needed" : latestRun ? "Synchronization is healthy" : "Ready for the first sync";
  const healthTone = !apiOnline || attention ? "danger" : configured ? "success" : "warning";
  const targetPolicies = (["google", "microsoft"] as TargetProvider[]).filter(target => target === "google" ? config.googleEnabled : config.microsoftEnabled).map(target => ({ target, policy: target === "google" ? config.syncPolicy : config.microsoftSyncPolicy }));
  return <>
    <section className={`health-banner ${healthTone}`}>
      <div className="health-orbit"><span>{healthTone === "success" ? "✓" : "!"}</span></div>
      <div><p className="eyebrow">Current sync health</p><h2>{healthTitle}</h2><p>{lastSync === "Never" ? "No completed sync has been recorded." : `Last completed ${lastSync}.`}</p></div>
      <div className="health-meta"><span className={`status-pill ${healthTone}`}><i /> {apiOnline ? health : "Offline"}</span><small>Every {config.interval} minutes</small></div>
    </section>

    <section className="target-health-grid" aria-label="Calendar target health">
      {(["google", "microsoft"] as TargetProvider[]).map(target => {
        const enabled = target === "google" ? config.googleEnabled : config.microsoftEnabled;
        const providerPeople = peopleByTarget[target];
        const providerCounts = targetCounts[target];
        const errors = providerPeople.length ? providerPeople.filter(person => person.syncEnabled && person.status === "Error").length : providerCounts.errors;
        const matched = providerPeople.length ? providerPeople.filter(person => person.status !== "Unmatched").length : Math.max(0, providerCounts.users - providerCounts.unmatched);
        return <article className={`target-health-card ${target} ${enabled ? "enabled" : "disabled"}`} key={target}>
          <div className="target-card-head"><span className={`provider-mark ${target}`}>{target === "google" ? "G" : "M"}</span><div><b>{TARGET_LABELS[target]}</b><small>{enabled ? errors ? "Enabled · attention required" : "Enabled · ready" : "Not enabled"}</small></div><span className={`status-pill ${enabled ? errors ? "danger" : "success" : "info"}`}><i />{enabled ? errors ? `${errors} error${errors === 1 ? "" : "s"}` : "Active" : "Off"}</span></div>
          <dl><div><dt>Matched</dt><dd>{matched}</dd></div><div><dt>Enabled</dt><dd>{providerPeople.length ? providerPeople.filter(person => person.status !== "Unmatched" && person.syncEnabled).length : providerCounts.enabled}</dd></div><div><dt>Managed items</dt><dd>{providerCounts.events}</dd></div></dl>
          <button className="text-button" onClick={() => enabled ? onOpenPeople(target) : onOpenConnections()}>{enabled ? `Manage ${TARGET_LABELS[target]} people` : `Configure ${TARGET_LABELS[target]}`} <span>→</span></button>
        </article>;
      })}
    </section>

    <section className="metric-grid" aria-label="Sync summary">
      <Metric label="Target accounts healthy" value={healthyUsers.toLocaleString()} detail={totalUsers ? `of ${totalUsers.toLocaleString()} enabled target accounts` : people.length ? "No accounts enabled" : "No accounts discovered"} delta={totalUsers ? `${(healthyUsers / totalUsers * 100).toFixed(1)}%` : "—"} />
      <Metric label="Delivered calendar items" value={(counts?.events ?? 0).toLocaleString()} detail="across all enabled targets" delta="One source item may be delivered twice" />
      <Metric label="Last run" value={latestRun?.duration ?? "—"} detail={latestRun ? `${latestRun.users.toLocaleString()} enabled people synced` : "No run recorded"} delta={latestRun?.status ?? "Waiting"} />
      <Metric label="Unmatched" value={String(unmatchedUsers)} detail="Target accounts without an active Schoolbox match" delta="Informational" />
    </section>

    <section className="dashboard-grid">
      <div className="panel activity-panel">
        <PanelHead title="Calendar activity" subtitle="Changes applied by the latest seven runs" action="View runs" onClick={() => onNavigate("runs")} />
        {activity.length ? <div className="chart-wrap" aria-label="Calendar changes bar chart">
          <div className="chart-key"><span><i className="created" />Created</span><span><i className="updated" />Updated</span></div>
          <div className="bar-chart">
            {activity.map((run, index) => <div className="bar-column" key={run.id} title={`${run.started}: ${run.created ?? 0} created, ${run.updated ?? 0} updated`}><div className="bars"><i className="bar-a" style={{ height: `${(run.created ?? 0) / activityMaximum * 100}%` }} /><i className="bar-b" style={{ height: `${(run.updated ?? 0) / activityMaximum * 100}%` }} /></div><span>{index + 1}</span></div>)}
          </div>
        </div> : <div className="empty-state compact"><b>{runsError ? "Activity could not be loaded" : "No activity yet"}</b><p>{runsError ? "Refresh after checking the server logs." : "Run the first sync to populate this chart."}</p></div>}
      </div>
      <div className="panel coverage-panel">
        <PanelHead title="Calendar coverage" subtitle="Current event policy by enabled target" />
        <div className="coverage-checks">{targetPolicies.map(({ target, policy }) => { const categories = EVENT_CATEGORIES.filter(category => policy.categories[category]); return <span key={target}><i>{target === "google" ? "G" : "M"}</i><b>{TARGET_LABELS[target]}</b><small>{categories.length} of {EVENT_CATEGORIES.length} categories · {policy.eventTypeMode === "all" ? "all exact types" : policy.eventTypeMode === "include" ? `${policy.eventTypes.length} allowed type(s)` : `${policy.eventTypes.length} excluded type(s)`}</small></span>; })}{targetPolicies.length === 0 && <span><i>!</i>No calendar target enabled</span>}</div>
        <div className="window-note"><span aria-hidden="true">↔</span><div><b>{Number(config.pastDays) + Number(config.futureDays)}-day rolling window</b><small>{config.pastDays} days back · {config.futureDays} days ahead</small></div></div>
      </div>
    </section>

    <section className="panel recent-panel">
      <PanelHead title="Recent runs" subtitle="The latest enabled-user calendar activity" action="All run history" onClick={() => onNavigate("runs")} />
      <div className="table-wrap"><table><thead><tr><th>Run</th><th>Started</th><th>Status</th><th>Synced</th><th>Changes</th><th>Duration</th><th><span className="sr-only">Open</span></th></tr></thead><tbody>{runs.slice(0, 4).map(run => <RunRow key={run.id} run={run} onClick={() => onSelectRun(run)} />)}{runs.length === 0 && <tr><td colSpan={7} className="table-empty">{runsError ? "Run history could not be loaded." : "No sync runs yet. Start a manual run when setup is complete."}</td></tr>}</tbody></table></div>
    </section>
  </>;
}

function Metric({ label, value, detail, delta, warning }: { label: string; value: string; detail: string; delta: string; warning?: boolean }) {
  return <div className="metric"><div className="metric-top"><span>{label}</span><span className={`metric-badge ${warning ? "warning" : ""}`}>{warning ? "!" : "↗"}</span></div><strong>{value}</strong><div><small>{detail}</small><b className={warning ? "text-warn" : ""}>{delta}</b></div></div>;
}

function PanelHead({ title, subtitle, action, onClick }: { title: string; subtitle: string; action?: string; onClick?: () => void }) {
  return <div className="panel-head"><div><h2>{title}</h2><p>{subtitle}</p></div>{action && <button className="text-button" onClick={onClick}>{action} <span aria-hidden="true">→</span></button>}</div>;
}

function RunRow({ run, onClick }: { run: Run; onClick: () => void }) {
  return <tr><td><button className="table-link" onClick={onClick}>{run.id}</button><small>{run.trigger}</small>{run.targets?.length ? <span className="run-target-chips">{run.targets.map(target => <i className={target.target} key={target.target}>{target.target === "google" ? "Google" : "Microsoft"}</i>)}</span> : null}</td><td>{run.started}</td><td><StatusPill status={run.status} />{run.status === "Running" && <small className="cell-detail">{run.phase.replaceAll("_", " ")}</small>}</td><td>{run.users.toLocaleString()}<small className="cell-detail">of {run.usersDiscovered.toLocaleString()} discovered target accounts</small></td><td>{run.changes.toLocaleString()}</td><td>{run.duration}</td><td><button className="row-open" onClick={onClick} aria-label={`Open ${run.id}`}>→</button></td></tr>;
}

function StatusPill({ status }: { status: Person["status"] | Run["status"] }) {
  const tone = status === "Succeeded" || status === "Synced" || status === "Syncing" ? "success" : status === "Failed" || status === "Error" ? "danger" : status === "Running" || status === "Unmatched" ? "info" : "warning";
  return <span className={`status-pill ${tone}`}><i />{status}</span>;
}

function diagnosticDate(value: string | null | undefined): string {
  if (!value) return "Not recorded";
  const parsed = Date.parse(value);
  return Number.isNaN(parsed) ? value : new Date(parsed).toLocaleString("en-AU");
}

function safeExternalUrl(value: string | null | undefined): string | null {
  if (!value) return null;
  try {
    const url = new URL(value);
    return url.protocol === "https:" || url.protocol === "http:" ? url.toString() : null;
  } catch {
    return null;
  }
}

function DiagnosticEventList({ events, empty = "No event detail was recorded." }: { events: DiagnosticEvent[]; empty?: string }) {
  if (events.length === 0) return <div className="diagnostic-empty">{empty}</div>;
  return <div className="diagnostic-events">{events.map(event =>
    <details className={`diagnostic-event action-${event.action ?? "managed"}`} key={`${event.runId ?? "current"}:${event.sourceKey}`}>
      <summary><div><b>{event.title || "Untitled event"}</b><small>{diagnosticDate(event.sourceStart)}{event.sourceEnd ? ` - ${diagnosticDate(event.sourceEnd)}` : ""}</small></div><span>{event.action ?? "managed"}</span></summary>
      <div className="diagnostic-event-body">
        {event.errorMessage && <p className="diagnostic-error"><b>Error</b>{event.errorMessage}</p>}
        {event.detail && <p>{event.detail}</p>}
        <dl>
          <div><dt>Event type</dt><dd>{event.eventType || "Not supplied"}</dd></div>
          <div><dt>Category</dt><dd>{event.category || "Not supplied"}</dd></div>
          <div><dt>Time format</dt><dd>{event.allDay ? "All day" : "Timed"}</dd></div>
          <div><dt>Destination</dt><dd>{event.destinationId || event.calendarId || "Primary / not recorded"}</dd></div>
          <div><dt>Source key</dt><dd><code>{event.sourceKey}</code></dd></div>
          <div><dt>{event.target === "microsoft" ? "Outlook event ID" : "Google event ID"}</dt><dd><code>{event.targetEventId || event.googleEventId || "Not created"}</code></dd></div>
          <div><dt>{event.target === "microsoft" ? "Outlook calendar ID" : "Google calendar ID"}</dt><dd><code>{event.calendarId || "Not recorded"}</code></dd></div>
          <div><dt>Last recorded</dt><dd>{diagnosticDate(event.recordedAt || event.updatedAt)}</dd></div>
        </dl>
        {event.location && <p><b>Location</b>{event.location}</p>}
        {event.author && <p><b>Author</b>{event.author}</p>}
        {event.description && <p className="diagnostic-description"><b>Description</b>{event.description}</p>}
        {safeExternalUrl(event.sourceUrl) && <a href={safeExternalUrl(event.sourceUrl) ?? undefined} target="_blank" rel="noreferrer">Open source in Schoolbox <span aria-hidden="true">↗</span></a>}
      </div>
    </details>,
  )}</div>;
}

type Config = {
  schoolboxUrl: string;
  schoolboxJwt: string;
  serviceAccountJson: string;
  adminEmail: string;
  interval: string;
  pastDays: string;
  futureDays: string;
  syncNewUsersByDefault: boolean;
  googleEnabled: boolean;
  syncNewGoogleUsersByDefault: boolean;
  googleCustomer: string;
  microsoftEnabled: boolean;
  microsoftTenantId: string;
  microsoftClientId: string;
  microsoftClientSecret: string;
  microsoftTestUserEmail: string;
  microsoftConsentGrantedAt: string;
  hasMicrosoftClientSecret: boolean;
  syncNewMicrosoftUsersByDefault: boolean;
  timezone: string;
  concurrency: string;
  discoveryTimeoutSeconds: string;
  userSyncTimeoutSeconds: string;
  runTimeoutMinutes: string;
  enabled: boolean;
  setupCompleted: boolean;
  schoolboxSetupCompleted: boolean;
  googleSetupCompleted: boolean;
  microsoftSetupCompleted: boolean;
  hasSchoolboxToken: boolean;
  hasGoogleServiceAccount: boolean;
  serviceAccountEmail: string;
  serviceAccountClientId: string;
  syncPolicy: SyncPolicy;
  microsoftSyncPolicy: SyncPolicy;
};

function mergeSavedConfig(current: Config, saved: Record<string, unknown>): Config {
  return {
    ...current,
    schoolboxUrl: String(saved.schoolboxBaseUrl ?? current.schoolboxUrl),
    schoolboxJwt: "",
    serviceAccountJson: "",
    microsoftClientSecret: "",
    adminEmail: String(saved.googleAdminEmail ?? current.adminEmail),
    googleCustomer: String(saved.googleCustomer ?? current.googleCustomer),
    googleEnabled: Boolean(saved.googleEnabled ?? current.googleEnabled),
    syncNewGoogleUsersByDefault: Boolean(saved.syncNewGoogleUsersByDefault ?? current.syncNewGoogleUsersByDefault),
    syncNewUsersByDefault: Boolean(saved.syncNewGoogleUsersByDefault ?? current.syncNewUsersByDefault),
    microsoftTenantId: String(saved.microsoftTenantId ?? current.microsoftTenantId),
    microsoftClientId: String(saved.microsoftClientId ?? current.microsoftClientId),
    microsoftTestUserEmail: String(saved.microsoftTestUserEmail ?? current.microsoftTestUserEmail),
    microsoftConsentGrantedAt: String(saved.microsoftConsentGrantedAt ?? current.microsoftConsentGrantedAt),
    microsoftEnabled: Boolean(saved.microsoftEnabled ?? current.microsoftEnabled),
    syncNewMicrosoftUsersByDefault: Boolean(saved.syncNewMicrosoftUsersByDefault ?? current.syncNewMicrosoftUsersByDefault),
    hasSchoolboxToken: Boolean(saved.hasSchoolboxToken ?? current.hasSchoolboxToken),
    hasGoogleServiceAccount: Boolean(saved.hasGoogleServiceAccount ?? current.hasGoogleServiceAccount),
    hasMicrosoftClientSecret: Boolean(saved.hasMicrosoftClientSecret ?? current.hasMicrosoftClientSecret),
    serviceAccountEmail: String(saved.serviceAccountEmail ?? current.serviceAccountEmail),
    serviceAccountClientId: String(saved.serviceAccountClientId ?? current.serviceAccountClientId),
    schoolboxSetupCompleted: Boolean(saved.schoolboxSetupCompleted ?? current.schoolboxSetupCompleted),
    googleSetupCompleted: Boolean(saved.googleSetupCompleted ?? current.googleSetupCompleted),
    microsoftSetupCompleted: Boolean(saved.microsoftSetupCompleted ?? current.microsoftSetupCompleted),
    setupCompleted: Boolean(saved.setupCompleted ?? current.setupCompleted),
    enabled: Boolean(saved.enabled ?? current.enabled),
  };
}

type DiscoveredEventType = {
  key: string;
  label: string;
  category: EventCategory;
  lastSeenAt: string;
};

function SetupWizard({ configured, track, setTrack, resume, config: savedConfig, setConfig: setSavedConfig, setConfigured, setNotice, changeView }: { configured: boolean; track: SetupTrack; setTrack: (track: SetupTrack) => void; resume?: SetupResume; config: Config; setConfig: React.Dispatch<React.SetStateAction<Config>>; setConfigured: (configured: boolean) => void; setNotice: (notice: Notice) => void; changeView: (view: View) => void }) {
  const applicationOrigin = useApplicationOrigin();
  const microsoftConsentCallback = `${applicationOrigin || "https://relay-host"}/api/auth/microsoft/admin-consent/callback`;
  const initialDraft = useMemo(() => ({
    ...savedConfig,
    ...(resume?.microsoftEnabled === undefined ? {} : { microsoftEnabled: resume.microsoftEnabled }),
    ...(resume?.syncNewMicrosoftUsersByDefault === undefined ? {} : { syncNewMicrosoftUsersByDefault: resume.syncNewMicrosoftUsersByDefault }),
  }), [resume, savedConfig]);
  const [config, setConfig] = useState<Config>(initialDraft);
  const [authoritative, setAuthoritative] = useState<Config>(savedConfig);
  const step = track === "schoolbox" ? 1 : track === "google" ? 3 : track === "microsoft" ? 4 : 0;
  const [testing, setTesting] = useState<"schoolbox" | "google" | "microsoft" | null>(null);
  const [schoolboxTested, setSchoolboxTested] = useState(resume?.schoolboxTested ?? config.schoolboxSetupCompleted);
  const [googleTested, setGoogleTested] = useState(resume?.googleTested ?? config.googleSetupCompleted);
  const [microsoftTested, setMicrosoftTested] = useState(resume?.microsoftTested ?? config.microsoftSetupCompleted);
  const [savingMicrosoftConsent, setSavingMicrosoftConsent] = useState(false);
  const clientId = useMemo(() => {
    try { const parsed = JSON.parse(config.serviceAccountJson || "{}"); return String(parsed.client_id ?? config.serviceAccountClientId); } catch { return config.serviceAccountClientId; }
  }, [config.serviceAccountJson, config.serviceAccountClientId]);

  const acceptScopedSave = (saved: Record<string, unknown>, desired: Partial<Config> = {}) => {
    setSavedConfig(current => mergeSavedConfig(current, saved));
    setAuthoritative(current => mergeSavedConfig(current, saved));
    setConfig(current => ({ ...mergeSavedConfig(current, saved), ...desired }));
    setConfigured(Boolean(saved.setupCompleted));
  };
  const leaveTrack = (next: SetupTrack = "hub") => {
    setConfig(authoritative);
    setSchoolboxTested(authoritative.schoolboxSetupCompleted);
    setGoogleTested(authoritative.googleSetupCompleted);
    setMicrosoftTested(authoritative.microsoftSetupCompleted);
    setTrack(next);
  };
  const saveTrackDraft = async (target: "schoolbox" | TargetProvider) => {
    const desired: Partial<Config> = target === "google"
      ? { googleEnabled: config.googleEnabled, syncNewGoogleUsersByDefault: config.syncNewGoogleUsersByDefault, syncNewUsersByDefault: config.syncNewGoogleUsersByDefault }
      : target === "microsoft"
        ? { microsoftEnabled: config.microsoftEnabled, syncNewMicrosoftUsersByDefault: config.syncNewMicrosoftUsersByDefault }
        : {};
    const payload = target === "schoolbox"
      ? {
          schoolboxBaseUrl: config.schoolboxUrl,
          ...(config.schoolboxJwt ? { schoolboxToken: config.schoolboxJwt } : {}),
        }
      : target === "google"
        ? {
            ...(config.serviceAccountJson ? { googleServiceAccountJson: config.serviceAccountJson } : {}),
            googleAdminEmail: config.adminEmail,
            googleCustomer: config.googleCustomer,
            syncNewGoogleUsersByDefault: config.syncNewGoogleUsersByDefault,
          }
        : {
            microsoftTenantId: config.microsoftTenantId,
            microsoftClientId: config.microsoftClientId,
            ...(config.microsoftClientSecret ? { microsoftClientSecret: config.microsoftClientSecret } : {}),
            microsoftTestUserEmail: config.microsoftTestUserEmail,
            syncNewMicrosoftUsersByDefault: config.syncNewMicrosoftUsersByDefault,
          };
    const saved = await fetchJson("/api/config", { method: "PUT", body: JSON.stringify(payload) });
    acceptScopedSave(saved, desired);
    return saved;
  };

  const testConnection = async (target: "schoolbox" | "google" | "microsoft") => {
    if (target === "schoolbox" && (!config.schoolboxUrl || (!config.schoolboxJwt && !config.hasSchoolboxToken))) {
      setNotice({ kind: "error", message: "Enter the Schoolbox URL and JWT before testing." }); return;
    }
    if (target === "google") {
      if (!config.adminEmail || (!config.serviceAccountJson && !config.hasGoogleServiceAccount)) {
        setNotice({ kind: "error", message: "Add service account credentials and a delegated admin email." }); return;
      }
      if (config.serviceAccountJson) {
        try { const parsed = JSON.parse(config.serviceAccountJson); if (!parsed.client_email || !parsed.private_key) throw new Error(); }
        catch { setNotice({ kind: "error", message: "Add valid service account JSON and a delegated admin email." }); return; }
      }
    }
    if (target === "microsoft" && (!config.microsoftTenantId || !config.microsoftClientId || (!config.microsoftClientSecret && !config.hasMicrosoftClientSecret) || !config.microsoftTestUserEmail)) {
      setNotice({ kind: "error", message: "Enter the Microsoft tenant ID, application ID, client secret and a test mailbox before testing." }); return;
    }
    setTesting(target);
    try {
      await saveTrackDraft(target);
      const data = await fetchJson("/api/diagnostics", { method: "POST", body: JSON.stringify({ target }) });
      if (target === "schoolbox") setSchoolboxTested(true); else if (target === "google") setGoogleTested(true); else {
        setMicrosoftTested(true);
        const consentStamp = String(data.microsoftConsentGrantedAt ?? config.microsoftConsentGrantedAt);
        setConfig(current => ({ ...current, microsoftConsentGrantedAt: consentStamp }));
        setSavedConfig(current => ({ ...current, microsoftConsentGrantedAt: consentStamp }));
        setAuthoritative(current => ({ ...current, microsoftConsentGrantedAt: consentStamp }));
      }
      setNotice({ kind: "success", message: `${String(data.message ?? `${target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target]} connection verified.`)} Choose Save setup to apply this target's activation choice.` });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Connection test failed.";
      setNotice({ kind: "error", message });
    } finally { setTesting(null); }
  };

  const copyScopes = async () => { await navigator.clipboard.writeText(SCOPES.join(",")); setNotice({ kind: "success", message: "Required scopes copied to clipboard." }); };
  const saveMicrosoftDraftAndOpenConsent = async () => {
    if (!config.microsoftTenantId || !config.microsoftClientId || (!config.microsoftClientSecret && !config.hasMicrosoftClientSecret) || !config.microsoftTestUserEmail) {
      setNotice({ kind: "error", message: "Enter the tenant ID, application ID, client secret and test mailbox before granting admin consent." });
      return;
    }
    setSavingMicrosoftConsent(true);
    try {
      await saveTrackDraft("microsoft");
      const consent = await fetchJson("/api/auth/microsoft/admin-consent/start", { method: "POST", body: "{}" });
      const consentUrl = typeof consent.url === "string" ? consent.url : "";
      if (!consentUrl) throw new Error("The Microsoft admin-consent URL was not returned.");
      window.sessionStorage.setItem("relay:microsoft-consent-progress", JSON.stringify({ track: "microsoft", schoolboxTested, googleTested, microsoftEnabled: config.microsoftEnabled, syncNewMicrosoftUsersByDefault: config.syncNewMicrosoftUsersByDefault }));
      window.location.assign(consentUrl);
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Microsoft credentials could not be saved before consent." });
      setSavingMicrosoftConsent(false);
    }
  };
  const completeTrack = async (target: Exclude<SetupTrack, "hub">) => {
    const verified = target === "schoolbox" ? schoolboxTested : target === "google" ? googleTested : microsoftTested;
    if (!verified) {
      setNotice({ kind: "error", message: `Verify ${target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target]} before completing this setup.` });
      return;
    }
    const label = target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target];
    try {
      const payload = target === "schoolbox"
        ? { schoolboxSetupCompleted: true }
        : target === "google"
          ? { googleSetupCompleted: true, googleEnabled: config.googleEnabled, syncNewGoogleUsersByDefault: config.syncNewGoogleUsersByDefault }
          : { microsoftSetupCompleted: true, microsoftEnabled: config.microsoftEnabled, syncNewMicrosoftUsersByDefault: config.syncNewMicrosoftUsersByDefault };
      const saved = await fetchJson("/api/config", { method: "PUT", body: JSON.stringify(payload) });
      acceptScopedSave(saved);
      setNotice({ kind: "success", message: `${label} setup saved. Other connections and scheduler settings were left unchanged.` });
      setTrack("hub");
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : `${label} setup could not be completed.` });
    }
  };

  if (track === "hub") return <section className="setup-hub">
    <div className={`setup-readiness ${configured ? "ready" : "pending"}`}><span>{configured ? "✓" : "i"}</span><div><p className="eyebrow">{configured ? "Relay operational" : "Connection readiness"}</p><h2>{configured ? "At least one calendar target is ready." : "Complete the source and one target."}</h2><p>Each connection has its own setup guide and completion state. Configure, revisit or disable a provider without stepping through the other provider.</p></div>{configured && <button type="button" className="button primary" onClick={() => changeView("dashboard")}>Go to overview <span>→</span></button>}</div>
    <div className="setup-hub-grid">
      <ConnectionSetupCard target="schoolbox" complete={config.schoolboxSetupCompleted} enabled detail={config.hasSchoolboxToken ? config.schoolboxUrl || "Credential stored" : "Schoolbox URL and JWT required"} onOpen={() => setTrack("schoolbox")} />
      <ConnectionSetupCard target="google" complete={config.googleSetupCompleted} enabled={config.googleEnabled} detail={config.hasGoogleServiceAccount ? config.adminEmail || "Service account stored" : "Service account and delegation required"} onOpen={() => setTrack("google")} />
      <ConnectionSetupCard target="microsoft" complete={config.microsoftSetupCompleted} enabled={config.microsoftEnabled} detail={config.hasMicrosoftClientSecret ? config.microsoftTenantId || "Entra credentials stored" : "Entra application and admin consent required"} onOpen={() => setTrack("microsoft")} />
    </div>
    <div className="setup-hub-footer"><div><b>Independent delivery</b><p>Google and Microsoft maintain separate directories, account switches, policies and managed events. Completing one target does not select or alter the other.</p></div><button type="button" className="button secondary" onClick={() => changeView("settings")}>Open sync settings</button></div>
  </section>;

  return <div className="setup-layout">
    <aside className="setup-steps setup-track-nav" aria-label="Connection setup">
      <p className="nav-label">Connections</p>
      <button onClick={() => leaveTrack("hub")}><span>←</span><div><b>All connections</b><small>Readiness overview</small></div></button>
      <button className={`${track === "schoolbox" ? "active" : ""} ${config.schoolboxSetupCompleted ? "done" : ""}`} onClick={() => leaveTrack("schoolbox")}><span>{config.schoolboxSetupCompleted ? "✓" : "S"}</span><div><b>Schoolbox</b><small>Shared source</small></div></button>
      <button className={`${track === "google" ? "active" : ""} ${config.googleSetupCompleted ? "done" : ""}`} onClick={() => leaveTrack("google")}><span>{config.googleSetupCompleted ? "✓" : "G"}</span><div><b>Google Workspace</b><small>{config.googleEnabled ? "Delivery enabled" : "Delivery disabled"}</small></div></button>
      <button className={`${track === "microsoft" ? "active" : ""} ${config.microsoftSetupCompleted ? "done" : ""}`} onClick={() => leaveTrack("microsoft")}><span>{config.microsoftSetupCompleted ? "✓" : "M"}</span><div><b>Microsoft 365</b><small>{config.microsoftEnabled ? "Delivery enabled" : "Delivery disabled"}</small></div></button>
      <div className="setup-help"><span>?</span><div><b>Provider setup</b><small>Keep this screen open while granting tenant-wide access in Google Admin or Microsoft Entra.</small><a href="https://learn.microsoft.com/entra/identity-platform/permissions-consent-overview" target="_blank" rel="noreferrer">Admin consent guide ↗</a></div></div>
    </aside>

    <section className="setup-card">
      <div className="setup-track-label"><button type="button" onClick={() => leaveTrack("hub")}>← All connections</button><span>Independent setup track</span></div>
      {step === 1 && <WizardSection eyebrow="Source connection" title="Connect your Schoolbox" intro="Relay reads each person’s timetable, events and due dates through the Schoolbox API. Your JWT is stored securely and never shown again.">
        <Field label="Schoolbox base URL" hint="The address your school uses to access Schoolbox."><div className="input-prefix"><span>https://</span><input value={config.schoolboxUrl.replace(/^https?:\/\//, "")} onChange={e => { setSchoolboxTested(false); setConfig(c => ({ ...c, schoolboxUrl: `https://${e.target.value}` })); }} placeholder="school.schoolbox.com.au" /></div></Field>
        <Field label="API JWT" hint={config.hasSchoolboxToken ? "A token is stored. Enter a value only to replace it." : "In Schoolbox Admin, edit the superuser, scroll to TOKENS, then choose Create token."}><input type="password" autoComplete="off" value={config.schoolboxJwt} onChange={e => { setSchoolboxTested(false); setConfig(c => ({ ...c, schoolboxJwt: e.target.value })); }} placeholder={config.hasSchoolboxToken ? "Stored securely" : "Paste your Schoolbox JWT"} /></Field>
        <div className="callout"><span>i</span><div><b>Give Relay read-only access</b><p>The token needs access to users, calendars, events and timetable data. It should not have permission to edit Schoolbox content.</p></div></div>
        <WizardActions><button className="button ghost" onClick={() => leaveTrack("hub")}>Cancel</button><button className="button secondary" onClick={() => void testConnection("schoolbox")} disabled={testing !== null}>{testing === "schoolbox" ? "Saving and testing…" : schoolboxTested ? "✓ Test saved connection again" : "Save draft and test"}</button><button className="button primary" onClick={() => void completeTrack("schoolbox")} disabled={!schoolboxTested}>Save Schoolbox setup</button></WizardActions>
      </WizardSection>}

      {step === 3 && <WizardSection eyebrow="Google Workspace" title="Connect Google Workspace" intro="Add the service account, grant domain-wide delegation, then verify directory and calendar access. This track does not alter Microsoft 365.">
        <>
          <div className="setup-target-options"><PolicyToggle checked={config.googleEnabled} onChange={enabled => setConfig(current => ({ ...current, googleEnabled: enabled }))} title="Enable Google Workspace delivery" detail="Directory discovery and Google Calendar synchronization can be enabled or disabled independently." /><PolicyToggle checked={config.syncNewGoogleUsersByDefault} onChange={enabled => setConfig(current => ({ ...current, syncNewGoogleUsersByDefault: enabled, syncNewUsersByDefault: enabled }))} title="Enable newly matched Google accounts" detail={config.syncNewGoogleUsersByDefault ? "Future matches start syncing immediately." : "Pilot-safe: future matches remain paused."} /></div>
          <ol className="instruction-list"><li><span>1</span><div><b>Create a delegated service account</b><p>Enable the Google Calendar API and Admin SDK API, create a service account, enable domain-wide delegation, and download a JSON key.</p><a href="https://console.cloud.google.com/iam-admin/serviceaccounts" target="_blank" rel="noreferrer">Open Google Cloud service accounts ↗</a></div></li><li><span>2</span><div><b>Grant domain-wide delegation</b><p>In Google Admin, add the numeric client ID and authorise the scopes shown below.</p><a href="https://admin.google.com/ac/owl/domainwidedelegation" target="_blank" rel="noreferrer">Open Google Admin ↗</a></div></li><li><span>3</span><div><b>Verify target access</b><p>Relay checks directory discovery and calendar permissions without changing user events.</p></div></li></ol>
          <Field label="Service account JSON" hint={config.hasGoogleServiceAccount ? "A service account is stored. Enter JSON only to replace it." : "Relay encrypts this credential at rest."}><textarea rows={6} value={config.serviceAccountJson} onChange={e => { setGoogleTested(false); setConfig(c => ({ ...c, serviceAccountJson: e.target.value })); }} placeholder={config.hasGoogleServiceAccount ? "Stored securely" : '{\n  "type": "service_account",\n  "client_id": "..."\n}'} /></Field>
          {clientId && <div className="detected-value"><span>✓</span><div><b>Service account recognised</b><small>OAuth client ID: {clientId}</small></div></div>}
          <Field label="Delegated administrator email" hint="Use a dedicated active administrator with directory access."><input type="email" value={config.adminEmail} onChange={e => { setGoogleTested(false); setConfig(c => ({ ...c, adminEmail: e.target.value })); }} placeholder="calendar-admin@example.edu" /></Field>
          <div className="delegation-box compact"><div className="delegation-number">2</div><div><h3>Authorise these scopes</h3><div className="scope-list">{SCOPES.map(scope => <code key={scope}>{scope}</code>)}</div><button type="button" className="text-button" onClick={() => void copyScopes()}>Copy all scopes <span>□</span></button></div></div>
          <div className={`validation-card ${googleTested ? "passed" : ""}`}><div className="validation-icon">{googleTested ? "✓" : "↻"}</div><div><h3>{googleTested ? "Google Workspace is verified" : "Verify Google access"}</h3><p>Checks service-account authentication, directory visibility and Calendar API access.</p></div><button type="button" className="button secondary" onClick={() => void testConnection("google")} disabled={testing !== null}>{testing === "google" ? "Checking…" : googleTested ? "Test again" : "Verify access"}</button></div>
        </>
        <WizardActions><button className="button ghost" onClick={() => leaveTrack("hub")}>Cancel</button><button className="button primary" onClick={() => void completeTrack("google")} disabled={!googleTested}>Save Google setup</button></WizardActions>
      </WizardSection>}

      {step === 4 && <WizardSection eyebrow="Microsoft 365" title="Connect Microsoft 365" intro="Create a single-tenant Entra application, grant application permissions with tenant admin consent, then verify one mailbox. This track does not alter Google Workspace.">
        <>
          <div className="setup-target-options"><PolicyToggle checked={config.microsoftEnabled} onChange={enabled => setConfig(current => ({ ...current, microsoftEnabled: enabled }))} title="Enable Microsoft 365 delivery" detail="Entra directory discovery and Outlook Calendar synchronization can be enabled or disabled independently." /><PolicyToggle checked={config.syncNewMicrosoftUsersByDefault} onChange={enabled => setConfig(current => ({ ...current, syncNewMicrosoftUsersByDefault: enabled }))} title="Enable newly matched Microsoft accounts" detail={config.syncNewMicrosoftUsersByDefault ? "Future matches start syncing immediately." : "Pilot-safe: future matches remain paused."} /></div>
          <ol className="instruction-list"><li><span>1</span><div><b>Register a single-tenant application</b><p>In Microsoft Entra, create an app registration for this organisation, add the exact Web redirect URI shown below, and copy its Directory (tenant) ID and Application (client) ID.</p><a href="https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade" target="_blank" rel="noreferrer">Open Microsoft Entra app registrations ↗</a></div></li><li><span>2</span><div><b>Add Microsoft Graph application permissions</b><p>Grant only <code>User.Read.All</code> and <code>Calendars.ReadWrite</code> as application permissions. Restrict mailbox access in Exchange where required.</p></div></li><li><span>3</span><div><b>Grant tenant admin consent</b><p>Save the credentials first, then use the interactive consent action. Relay verifies permissions against only the test mailbox below.</p></div></li></ol>
          <Field label="Web redirect URI" hint="Register this exact URI under Authentication → Web in the Entra app before starting consent."><CopyBox value={microsoftConsentCallback} onCopy={() => void navigator.clipboard.writeText(microsoftConsentCallback)} /></Field>
          <div className="form-grid two"><Field label="Directory (tenant) ID"><input required value={config.microsoftTenantId} onChange={event => { setMicrosoftTested(false); setConfig(current => ({ ...current, microsoftTenantId: event.target.value })); }} placeholder="00000000-0000-0000-0000-000000000000" /></Field><Field label="Application (client) ID"><input required value={config.microsoftClientId} onChange={event => { setMicrosoftTested(false); setConfig(current => ({ ...current, microsoftClientId: event.target.value })); }} placeholder="00000000-0000-0000-0000-000000000000" /></Field></div>
          <Field label="Client secret" hint={config.hasMicrosoftClientSecret ? "A secret is stored. Leave blank to retain it, or enter a replacement." : "Paste the secret value, not its secret ID. Relay encrypts it at rest."}><input type="password" autoComplete="off" value={config.microsoftClientSecret} onChange={event => { setMicrosoftTested(false); setConfig(current => ({ ...current, microsoftClientSecret: event.target.value })); }} placeholder={config.hasMicrosoftClientSecret ? "Stored securely" : "Microsoft Entra client secret value"} /></Field>
          <Field label="Test mailbox" hint="Required. An active mailbox used only to verify calendar permissions."><input required type="email" value={config.microsoftTestUserEmail} onChange={event => { setMicrosoftTested(false); setConfig(current => ({ ...current, microsoftTestUserEmail: event.target.value })); }} placeholder="relay-test@example.edu" /></Field>
          <div className="consent-actions"><button type="button" className="button ghost" onClick={() => void saveMicrosoftDraftAndOpenConsent()} disabled={savingMicrosoftConsent || !config.microsoftTenantId || !config.microsoftClientId || (!config.microsoftClientSecret && !config.hasMicrosoftClientSecret) || !config.microsoftTestUserEmail.trim()}>{savingMicrosoftConsent ? "Saving…" : config.hasMicrosoftClientSecret ? "Renew admin consent" : "Save credentials and grant consent"} <span>↗</span></button><small>{config.microsoftConsentGrantedAt ? `Admin consent and Microsoft Graph access verified ${diagnosticDate(config.microsoftConsentGrantedAt)}. The target remains disabled until setup is completed.` : "Credentials are saved with Microsoft delivery disabled, then Microsoft sign-in opens for a tenant administrator."}</small></div>
          <div className={`validation-card ${microsoftTested ? "passed" : ""}`}><div className="validation-icon">{microsoftTested ? "✓" : "↻"}</div><div><h3>{microsoftTested ? "Microsoft 365 is verified" : "Verify Microsoft Graph access"}</h3><p>Checks application authentication and Entra directory discovery, then creates and immediately deletes a temporary secondary calendar in the test mailbox as a write probe.</p></div><button type="button" className="button secondary" onClick={() => void testConnection("microsoft")} disabled={testing !== null || !config.microsoftTestUserEmail.trim()}>{testing === "microsoft" ? "Checking…" : microsoftTested ? "Test again" : "Verify access"}</button></div>
        </>
        <WizardActions><button className="button ghost" onClick={() => leaveTrack("hub")}>Cancel</button><button className="button primary" onClick={() => void completeTrack("microsoft")} disabled={!microsoftTested}>Save Microsoft setup</button></WizardActions>
      </WizardSection>}

    </section>
  </div>;
}

function WizardSection({ eyebrow, title, intro, children }: { eyebrow: string; title: string; intro: string; children: React.ReactNode }) { return <><p className="eyebrow">{eyebrow}</p><h2 className="setup-title">{title}</h2><p className="setup-intro">{intro}</p><div className="setup-content">{children}</div></>; }
function WizardActions({ children }: { children: React.ReactNode }) { return <div className="wizard-actions">{children}</div>; }
function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) { return <label className="field"><span>{label}</span>{children}{hint && <small>{hint}</small>}</label>; }
function CopyBox({ value, onCopy }: { value: string; onCopy: () => void }) { const [done, setDone] = useState(false); return <div className="copy-box"><code>{value}</code><button type="button" onClick={() => { onCopy(); setDone(true); window.setTimeout(() => setDone(false), 1500); }}>{done ? "Copied" : "Copy"}</button></div>; }
function ConnectionSetupCard({ target, complete, enabled, detail, onOpen }: { target: "schoolbox" | TargetProvider; complete: boolean; enabled: boolean; detail: string; onOpen: () => void }) {
  const label = target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target];
  return <article className={`setup-connection-card ${complete ? "complete" : "pending"}`}><div className="setup-connection-head"><span className={target === "schoolbox" ? "connection-logo" : `provider-mark ${target}`}>{target === "schoolbox" ? "S" : target === "google" ? "G" : "M"}</span><span className={`status-pill ${complete ? "success" : "warning"}`}><i />{complete ? "Setup complete" : "Setup required"}</span></div><h3>{label}</h3><p>{detail}</p><div className="setup-connection-meta"><span>{target === "schoolbox" ? "Shared event source" : enabled ? "Delivery enabled" : "Delivery disabled"}</span><span>{complete ? "Verified connection" : "Not yet verified"}</span></div><button type="button" className={complete ? "button secondary" : "button primary"} onClick={onOpen}>{complete ? "Review setup" : `Set up ${label}`} <span>→</span></button></article>;
}

function PeoplePage({ initialTarget, peopleByTarget, setPeopleByTarget, targetCounts, loadErrors, configuredTargets, canConfigure, setNotice, onOpenConnections }: {
  initialTarget?: TargetProvider;
  peopleByTarget: Record<TargetProvider, Person[]>;
  setPeopleByTarget: React.Dispatch<React.SetStateAction<Record<TargetProvider, Person[]>>>;
  targetCounts: Record<TargetProvider, TargetCounts>;
  loadErrors: Record<TargetProvider, boolean>;
  configuredTargets: Record<TargetProvider, boolean>;
  canConfigure: boolean;
  setNotice: (notice: Notice) => void;
  onOpenConnections: () => void;
}) {
  const [target, setTarget] = useState<TargetProvider>(initialTarget ?? (configuredTargets.google ? "google" : "microsoft"));
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("All statuses");
  const [coverageFilter, setCoverageFilter] = useState("All coverage");
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [selectedPerson, setSelectedPerson] = useState<Person | null>(null);
  const [userDetail, setUserDetail] = useState<UserDetailPayload | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailError, setDetailError] = useState("");
  const [page, setPage] = useState(0);
  const pageSize = 100;
  const selectAllRef = useRef<HTMLInputElement>(null);
  const people = peopleByTarget[target];
  const counts = targetCounts[target];
  const loadError = loadErrors[target];
  const setPeople = (updater: React.SetStateAction<Person[]>) => setPeopleByTarget(current => ({
    ...current,
    [target]: typeof updater === "function" ? updater(current[target]) : updater,
  }));
  const filtered = useMemo(() => people.filter(person => {
    const statusMatches = statusFilter === "All statuses" || person.status === statusFilter;
    const matchable = person.status !== "Unmatched";
    const coverageMatches = coverageFilter === "All coverage" || (coverageFilter === "Enabled" ? matchable && person.syncEnabled : matchable && !person.syncEnabled);
    return statusMatches && coverageMatches && `${person.name} ${person.schoolboxEmail} ${person.targetEmail}`.toLowerCase().includes(query.toLowerCase());
  }), [people, query, statusFilter, coverageFilter]);
  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const pageIndex = Math.min(page, pageCount - 1);
  const visible = filtered.slice(pageIndex * pageSize, (pageIndex + 1) * pageSize);
  const selectableVisible = visible.filter(person => person.status !== "Unmatched");
  const selectedIds = selectableVisible.filter(person => selected.has(person.id)).map(person => person.id);
  const selectedVisible = selectedIds.length;
  const enabledCount = people.length ? people.filter(person => person.status !== "Unmatched" && person.syncEnabled).length : counts.enabled;
  const pausedCount = people.length ? people.filter(person => person.status !== "Unmatched" && !person.syncEnabled).length : counts.disabled;

  const switchTarget = (nextTarget: TargetProvider) => {
    setTarget(nextTarget);
    setQuery("");
    setStatusFilter("All statuses");
    setCoverageFilter("All coverage");
    setSelected(new Set());
    setSelectedPerson(null);
    setPage(0);
  };

  const openPersonDetail = async (person: Person) => {
    setSelectedPerson(person);
    setUserDetail(null);
    setDetailError("");
    setDetailLoading(true);
    try {
      const payload = await fetchJson(`/api/user-details?target=${person.target}&userId=${encodeURIComponent(person.id)}`);
      setUserDetail(payload as unknown as UserDetailPayload);
    } catch (error) {
      setDetailError(error instanceof Error ? error.message : "User diagnostics could not be loaded.");
    } finally {
      setDetailLoading(false);
    }
  };

  useEffect(() => {
    if (selectAllRef.current) selectAllRef.current.indeterminate = selectedVisible > 0 && selectedVisible < selectableVisible.length;
  }, [selectableVisible.length, selectedVisible]);

  const selectVisible = (checked: boolean) => {
    setSelected(checked ? new Set(selectableVisible.map(person => person.id)) : new Set());
  };
  const selectOne = (id: string, checked: boolean) => {
    setSelected(current => {
      const next = new Set(current);
      if (checked) next.add(id); else next.delete(id);
      return next;
    });
  };
  const updateCoverage = async (ids: string[], syncEnabled: boolean) => {
    if (!ids.length || busy) return;
    setBusy(true);
    try {
      const payload = await fetchJson("/api/users", {
        method: "PATCH",
        body: JSON.stringify({ target, userIds: ids, syncEnabled }),
      });
      const updated = Number(payload.updated ?? ids.length);
      const changed = new Set(ids);
      setPeople(current => current.map(person => changed.has(person.id) ? { ...person, syncEnabled } : person));
      setSelected(new Set());
      setNotice({
        kind: "success",
        message: `${TARGET_LABELS[target]} calendar sync ${syncEnabled ? "enabled" : "paused"} for ${updated} ${updated === 1 ? "person" : "people"}. The change applies on the next run.`,
      });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "User sync coverage could not be updated." });
    } finally {
      setBusy(false);
    }
  };
  const cleanupManagedEvents = async (person: Person, deleteCalendars = false) => {
    if (busy || (person.eventCount === 0 && (!deleteCalendars || person.calendarCount === 0))) return;
    const confirmed = window.confirm(deleteCalendars
      ? `Pause ${TARGET_LABELS[person.target]} sync for ${person.name}, remove ${person.eventCount} Relay-managed event(s), then permanently delete ${person.calendarCount} Relay-created secondary calendar(s)? Every event in those calendars, including manually added events, will be permanently deleted. The primary calendar and unrelated calendars will not be touched.`
      : `Pause ${TARGET_LABELS[person.target]} sync for ${person.name} and remove ${person.eventCount} Relay-managed ${person.eventCount === 1 ? "event" : "events"} from their ${TARGET_CALENDAR_LABELS[person.target]}? Other calendar entries and the other provider will not be touched.`);
    if (!confirmed) return;
    setBusy(true);
    try {
      const payload = await fetchJson("/api/users", {
        method: "DELETE",
        body: JSON.stringify({ target: person.target, userId: person.id, deleteCalendars }),
      });
      const deleted = Number(payload.deleted ?? 0);
      const alreadyMissing = Number(payload.alreadyMissing ?? 0);
      const remaining = Number(payload.remaining ?? 0);
      const calendarsDeleted = Number(payload.calendarsDeleted ?? 0);
      const calendarsAlreadyMissing = Number(payload.calendarsAlreadyMissing ?? 0);
      const calendarsRemaining = Number(payload.calendarsRemaining ?? person.calendarCount);
      const cleanupError = typeof payload.error === "string" ? payload.error : null;
      setPeople(current => current.map(row => row.id === person.id ? {
        ...row,
        syncEnabled: false,
        eventCount: remaining,
        calendarCount: calendarsRemaining,
        status: remaining > 0 || cleanupError ? "Error" : row.status === "Unmatched" ? "Unmatched" : "Pending",
      } : row));
      if (remaining > 0 || (deleteCalendars && calendarsRemaining > 0) || cleanupError) {
        setNotice({ kind: "error", message: `Cleanup paused this ${TARGET_LABELS[person.target]} account and removed ${deleted} event(s) and ${calendarsDeleted} calendar(s), but ${remaining} event(s) and ${calendarsRemaining} calendar(s) remain. Retry after checking provider access.` });
      } else {
        const missingNote = alreadyMissing > 0 ? ` ${alreadyMissing} tracked event(s) were already absent.` : "";
        const calendarNote = deleteCalendars
          ? ` ${calendarsDeleted} Relay-created calendar(s) deleted.${calendarsAlreadyMissing > 0 ? ` ${calendarsAlreadyMissing} tracked calendar(s) were already absent.` : ""}`
          : "";
        setNotice({ kind: "success", message: `Calendar sync paused and ${deleted} Relay-managed event(s) removed.${missingNote}${calendarNote}` });
      }
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Relay-managed calendar data could not be removed." });
    } finally {
      setBusy(false);
    }
  };
  const exportCsv = () => {
    const escape = (value: string) => `"${value.replaceAll('"', '""')}"`;
    const rows = [["Target", "Name", "Schoolbox email", `${TARGET_LABELS[target]} email`, "Role", "Calendar sync", "Relay-managed events", "Relay-created calendars", "Status", "Last sync"], ...filtered.map(person => [TARGET_LABELS[target], person.name, person.schoolboxEmail, person.targetEmail, person.role, person.status === "Unmatched" ? "Unavailable" : person.syncEnabled ? "Enabled" : "Paused", String(person.eventCount), String(person.calendarCount), person.status, person.lastSync])];
    const url = URL.createObjectURL(new Blob([rows.map(row => row.map(escape).join(",")).join("\n")], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = `relay-${target}-user-mappings.csv`;
    link.click();
    URL.revokeObjectURL(url);
  };
  return <>
    <div className="provider-tabs" role="tablist" aria-label="Calendar target directory">
      {(["google", "microsoft"] as TargetProvider[]).map(provider => <button type="button" role="tab" aria-selected={target === provider} className={target === provider ? "active" : ""} onClick={() => switchTarget(provider)} key={provider}><span className={`provider-mark ${provider}`}>{provider === "google" ? "G" : "M"}</span><span><b>{TARGET_LABELS[provider]}</b><small>{configuredTargets[provider] ? `${targetCounts[provider].enabled} enabled · ${targetCounts[provider].users} discovered` : "Not enabled"}</small></span></button>)}
    </div>
    {!configuredTargets[target] ? <section className="panel target-empty-state"><span className={`provider-mark ${target}`}>{target === "google" ? "G" : "M"}</span><div><p className="eyebrow">Calendar target not enabled</p><h2>Enable {TARGET_LABELS[target]} to discover and manage its people.</h2><p>This target has its own directory, user selections, calendar mappings and cleanup controls. Enabling it does not alter the other target.</p></div>{canConfigure && <button className="button primary" type="button" onClick={onOpenConnections}>Open Settings → Connections</button>}</section> : <>
    <section className="people-summary"><div><span className="summary-icon green">#</span><p><b>{people.length || counts.users || 0}</b><small>Discovered</small></p></div><div><span className="summary-icon green">✓</span><p><b>{enabledCount}</b><small>Enabled</small></p></div><div><span className="summary-icon amber">Ⅱ</span><p><b>{pausedCount}</b><small>Paused</small></p></div><div><span className="summary-icon blue">○</span><p><b>{people.length ? people.filter(p => p.status === "Unmatched").length : counts.unmatched}</b><small>Unmatched</small></p></div></section>
    <section className="panel people-panel" aria-busy={busy}><div className="people-tools"><div className="search-box"><span aria-hidden="true">⌕</span><input value={query} onChange={e => { setQuery(e.target.value); setPage(0); setSelected(new Set()); }} placeholder="Search people or email…" aria-label="Search people" /></div><select value={coverageFilter} onChange={e => { setCoverageFilter(e.target.value); setPage(0); setSelected(new Set()); }} aria-label="Filter calendar sync coverage"><option>All coverage</option><option>Enabled</option><option>Paused</option></select><select value={statusFilter} onChange={e => { setStatusFilter(e.target.value); setPage(0); setSelected(new Set()); }} aria-label="Filter sync status"><option>All statuses</option><option>Synced</option><option>Syncing</option><option>Pending</option><option>Unmatched</option><option>Error</option></select><button className="button ghost" onClick={exportCsv}>Export CSV</button></div>
      {canConfigure && selectedVisible > 0 && <div className="people-bulk" role="status"><b>{selectedVisible} selected</b><span>Bulk changes apply only to the selected visible users.</span><button className="button secondary" onClick={() => void updateCoverage(selectedIds, true)} disabled={busy}>Enable selected</button><button className="button ghost" onClick={() => void updateCoverage(selectedIds, false)} disabled={busy}>Pause selected</button></div>}
      <div className="coverage-note"><span>i</span><p><b>Select a person’s name for detailed settings and diagnostics.</b> Individual event exclusions live in the person drawer, keeping this overview focused on sync coverage.</p></div>
      <div className="table-wrap">
        <table className="people-table">
          <caption className="sr-only">Discovered {TARGET_LABELS[target]} users and their Schoolbox calendar sync coverage</caption>
          <thead><tr>
            {canConfigure && <th scope="col" className="selection-column"><input ref={selectAllRef} type="checkbox" checked={selectableVisible.length > 0 && selectedVisible === selectableVisible.length} onChange={event => selectVisible(event.target.checked)} aria-label="Select visible matched users" disabled={busy || selectableVisible.length === 0} /></th>}
            <th scope="col">Person</th><th scope="col">Schoolbox identity</th><th scope="col">{TARGET_LABELS[target]}</th><th scope="col">Role</th><th scope="col">Calendar sync</th><th scope="col">Relay data</th><th scope="col">Status</th><th scope="col">Last sync</th>
          </tr></thead>
          <tbody>{visible.map(person => <tr key={person.id}>
            {canConfigure && <td className="selection-column"><input type="checkbox" checked={person.status !== "Unmatched" && selected.has(person.id)} onChange={event => selectOne(person.id, event.target.checked)} aria-label={person.status === "Unmatched" ? `${person.name} cannot be selected until matched` : `Select ${person.name}`} disabled={busy || person.status === "Unmatched"} /></td>}
            <th scope="row" className="person-row-header"><div className="person-cell"><span className="person-avatar">{person.name.split(" ").map(part => part[0]).join("").slice(0, 2)}</span><div><button type="button" className="person-link" onClick={() => void openPersonDetail(person)}>{person.name}</button><small>{person.id}</small></div></div></th>
            <td>{person.schoolboxEmail}</td>
            <td className={person.targetEmail === "—" ? "muted" : ""}>{person.targetEmail}</td>
            <td>{person.role}</td>
            <td>{person.status === "Unmatched" ? <span className="coverage-state unavailable" title="A Schoolbox identity must be discovered before calendar sync can be enabled">Unavailable</span> : canConfigure ? <label className="sync-switch"><input type="checkbox" checked={person.syncEnabled} onChange={event => void updateCoverage([person.id], event.target.checked)} disabled={busy} aria-label={`Sync calendar for ${person.name}`} /><span aria-hidden="true" /><b>{person.syncEnabled ? "Enabled" : "Paused"}</b></label> : <span className={`coverage-state ${person.syncEnabled ? "enabled" : "paused"}`}>{person.syncEnabled ? "Enabled" : "Paused"}</span>}</td>
            <td><div className="managed-events-cell"><span><b>{person.eventCount}</b> event(s)</span><span><b>{person.calendarCount}</b> calendar(s)</span>{person.hasCustomExclusions && <span className="custom-policy-marker">Custom exclusions</span>}{canConfigure && <button type="button" onClick={() => void cleanupManagedEvents(person)} disabled={busy || person.eventCount === 0} title={person.eventCount === 0 ? "No Relay-managed events to remove" : `Pause this account and remove only Relay-managed ${TARGET_CALENDAR_LABELS[target]} events`}>Remove Relay events</button>}{canConfigure && person.calendarCount > 0 && <button type="button" onClick={() => void cleanupManagedEvents(person, true)} disabled={busy} title={`Pause this account, remove Relay-managed events, and delete Relay-created ${TARGET_CALENDAR_LABELS[target]} calendars`}>Delete Relay calendars</button>}</div></td>
            <td><StatusPill status={person.status} /></td><td>{person.lastSync}</td>
          </tr>)}</tbody>
        </table>
        {filtered.length === 0 && <div className="empty-state"><b>{loadError ? `${TARGET_LABELS[target]} people could not be loaded` : people.length === 0 ? `No ${TARGET_LABELS[target]} accounts discovered yet` : "No people found"}</b><p>{loadError ? "Refresh after checking the provider connection and server logs." : people.length === 0 ? `Run a sync to discover ${TARGET_LABELS[target]} accounts. If this target's new-account default is paused, discovery will not write calendar events.` : "Try different search or filter options."}</p></div>}
      </div>
      <div className="table-footer"><span>{filtered.length ? `Showing ${pageIndex * pageSize + 1}–${Math.min((pageIndex + 1) * pageSize, filtered.length)} of ${filtered.length} matching people` : `0 of ${people.length} people`}</span><div><button onClick={() => { setPage(Math.max(0, pageIndex - 1)); setSelected(new Set()); }} disabled={pageIndex === 0}>Previous</button><span>Page {pageIndex + 1} of {pageCount}</span><button onClick={() => { setPage(Math.min(pageCount - 1, pageIndex + 1)); setSelected(new Set()); }} disabled={pageIndex >= pageCount - 1}>Next</button></div></div>
    </section>
    {selectedPerson && <UserDiagnosticsDrawer key={`${selectedPerson.target}:${selectedPerson.id}`}
      person={selectedPerson}
      detail={userDetail}
      loading={detailLoading}
      error={detailError}
      canConfigure={canConfigure}
      onClose={() => setSelectedPerson(null)}
      onRetry={() => void openPersonDetail(selectedPerson)}
      onPreferencesSaved={exclusions => {
        setUserDetail(current => current ? { ...current, exclusions } : current);
        setPeople(current => current.map(candidate => candidate.id === selectedPerson.id
          ? { ...candidate, hasCustomExclusions: exclusions.categories.length + exclusions.eventTypes.length > 0 }
          : candidate));
        setNotice({ kind: "success", message: "Individual event exclusions saved. They apply to this Schoolbox person on every enabled delivery target." });
      }}
    />}
    </>}
  </>;
}

function UserExclusionEditor({ person, detail, canConfigure, onSaved }: {
  person: Person;
  detail: UserDetailPayload;
  canConfigure: boolean;
  onSaved: (exclusions: UserEventExclusions) => void;
}) {
  const [categories, setCategories] = useState<EventCategory[]>(detail.exclusions.categories);
  const [eventTypes, setEventTypes] = useState<string[]>(detail.exclusions.eventTypes);
  const [query, setQuery] = useState("");
  const [manualType, setManualType] = useState("");
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");

  const catalog = useMemo(() => {
    const entries = new Map(detail.eventTypes.map(entry => [entry.key, entry]));
    for (const label of eventTypes) {
      const key = eventTypeKey(label);
      if (key && !entries.has(key)) entries.set(key, {
        key,
        label,
        category: "other",
        lastSeenAt: "",
      });
    }
    return [...entries.values()].sort((a, b) => a.label.localeCompare(b.label, "en-AU"));
  }, [detail.eventTypes, eventTypes]);
  const excludedTypeKeys = new Set(eventTypes.map(eventTypeKey));
  const visibleTypes = catalog.filter(entry =>
    `${entry.label} ${EVENT_CATEGORY_COPY[entry.category][0]}`.toLocaleLowerCase("en-AU")
      .includes(query.trim().toLocaleLowerCase("en-AU")),
  );
  const savedCategories = [...detail.exclusions.categories].sort().join("|");
  const draftCategories = [...categories].sort().join("|");
  const savedTypes = detail.exclusions.eventTypes.map(eventTypeKey).sort().join("|");
  const draftTypes = eventTypes.map(eventTypeKey).sort().join("|");
  const dirty = savedCategories !== draftCategories || savedTypes !== draftTypes;
  const exclusionCount = categories.length + eventTypes.length;

  const globalTypeCoverage = (entry: DiscoveredEventType) => {
    if (!entry.lastSeenAt) return "Saved exclusion · type not currently detected";
    const override = detail.globalPolicy.eventTypeOverrides[entry.key];
    if (override?.enabled === true) return "Included globally by exact rule";
    if (override?.enabled === false) return "Already excluded globally";
    if (!detail.globalPolicy.categories[entry.category]) return "Already excluded globally by category";
    if (detail.globalPolicy.eventTypeMode === "all") return "Type included globally";
    const listed = detail.globalPolicy.eventTypes.some(type => eventTypeKey(type) === entry.key);
    if (detail.globalPolicy.eventTypeMode === "include") return listed ? "Type included globally" : "Already excluded globally";
    return listed ? "Already excluded globally" : "Type included globally";
  };
  const toggleCategory = (category: EventCategory, excluded: boolean) => {
    setCategories(current => excluded
      ? [...new Set([...current, category])]
      : current.filter(candidate => candidate !== category));
  };
  const toggleType = (label: string, excluded: boolean) => {
    const key = eventTypeKey(label);
    setEventTypes(current => excluded
      ? current.some(candidate => eventTypeKey(candidate) === key) ? current : [...current, label]
      : current.filter(candidate => eventTypeKey(candidate) !== key));
  };
  const addManualType = (event: FormEvent) => {
    event.preventDefault();
    const label = normalizeEventTypeLabel(manualType);
    if (!label) return;
    toggleType(label, true);
    setManualType("");
    setQuery(label);
  };
  const save = async () => {
    if (!dirty || saving) return;
    setSaving(true);
    setSaveError("");
    try {
      const payload = await fetchJson("/api/user-details", {
        method: "PATCH",
        body: JSON.stringify({
          target: person.target,
          userId: person.id,
          excludedCategories: categories,
          excludedEventTypes: eventTypes,
        }),
      });
      onSaved(payload.exclusions as UserEventExclusions);
    } catch (saveFailure) {
      setSaveError(saveFailure instanceof Error ? saveFailure.message : "Individual exclusions could not be saved.");
    } finally {
      setSaving(false);
    }
  };

  return <div className="user-exclusion-editor">
    <div className="user-exclusion-intro"><span>−</span><div><b>Exclude events for this Schoolbox person across all targets</b><p>These exclusions are shared across Google Workspace and Microsoft 365 because they describe the source person. The provider badge above scopes only the managed items, diagnostics and cleanup shown in this drawer.</p></div></div>
    <div className="user-exclusion-summary"><b>{exclusionCount}</b><span>individual exclusion{exclusionCount === 1 ? "" : "s"}</span><small>{detail.exclusions.updatedAt ? `Last saved ${diagnosticDate(detail.exclusions.updatedAt)}` : "Using organisation defaults"}</small></div>

    <h3>Category exclusions</h3>
    <div className="user-category-exclusions">{EVENT_CATEGORIES.map(category => <label key={category}>
      <input type="checkbox" checked={categories.includes(category)} onChange={event => toggleCategory(category, event.target.checked)} disabled={!canConfigure || saving} />
      <span><b>{EVENT_CATEGORY_COPY[category][0]}</b><small>{detail.globalPolicy.categories[category] ? "Category baseline is included globally" : "Category baseline is excluded globally"}</small></span>
    </label>)}</div>

    <div className="user-type-heading"><div><h3>Exact Schoolbox type exclusions</h3><p>Use an exact type when this person still needs other events from the same category.</p></div>{canConfigure && <input value={query} onChange={event => setQuery(event.target.value)} placeholder="Search detected types…" aria-label="Search detected Schoolbox event types" />}</div>
    <div className="user-type-exclusions">
      {visibleTypes.map(entry => <label key={entry.key}>
        <input type="checkbox" checked={excludedTypeKeys.has(entry.key)} onChange={event => toggleType(entry.label, event.target.checked)} disabled={!canConfigure || saving} />
        <span><b>{entry.label}</b><small>{EVENT_CATEGORY_COPY[entry.category][0]} · {categories.includes(entry.category) ? "Covered by category exclusion" : globalTypeCoverage(entry)}</small></span>
      </label>)}
      {visibleTypes.length === 0 && <div className="diagnostic-empty">No detected event types match this search.</div>}
    </div>
    {canConfigure && <form className="manual-type-exclusion" onSubmit={addManualType}><Field label="Exclude an exact type not yet detected" hint="The label must exactly match the Schoolbox event type."><input maxLength={120} value={manualType} onChange={event => setManualType(event.target.value)} placeholder="Exact Schoolbox type label" /></Field><button className="button ghost" type="submit" disabled={!manualType.trim() || saving}>Add exclusion</button></form>}

    {saveError && <div className="diagnostic-error-box"><b>Could not save exclusions</b><p>{saveError}</p></div>}
    {canConfigure && <div className="user-exclusion-actions"><div><b>{dirty ? "Unsaved changes" : "Saved"}</b><small>Changes apply on this Schoolbox person’s next enabled sync in every target. Existing Relay events follow each target’s reconciliation setting.</small></div><button className="button ghost" type="button" onClick={() => { setCategories([]); setEventTypes([]); }} disabled={saving || exclusionCount === 0}>Clear all</button><button className="button primary" type="button" onClick={() => void save()} disabled={saving || !dirty}>{saving ? "Saving…" : "Save exclusions"}</button></div>}
  </div>;
}

function UserDiagnosticsDrawer({ person, detail, loading, error, canConfigure, onClose, onRetry, onPreferencesSaved }: {
  person: Person;
  detail: UserDetailPayload | null;
  loading: boolean;
  error: string;
  canConfigure: boolean;
  onClose: () => void;
  onRetry: () => void;
  onPreferencesSaved: (exclusions: UserEventExclusions) => void;
}) {
  const [section, setSection] = useState<"overview" | "exclusions">("overview");
  return <div className="detail-backdrop" role="presentation" onMouseDown={onClose}><aside className="admin-detail-drawer" role="dialog" aria-modal="true" aria-label={`Diagnostics for ${person.name}`} onMouseDown={event => event.stopPropagation()}>
    <div className="drawer-head"><div><p className="eyebrow">Person diagnostics · {TARGET_LABELS[person.target]}</p><h2>{person.name}</h2><small>{person.targetEmail}</small></div><button onClick={onClose} aria-label="Close person diagnostics">×</button></div>
    {loading && <div className="diagnostic-empty">Loading identities, calendars, events, and run history…</div>}
    {error && <div className="diagnostic-error-box"><b>Diagnostics unavailable</b><p>{error}</p><button className="button secondary" onClick={onRetry}>Try again</button></div>}
    {detail && <>
      <div className="drawer-summary"><span><small>Status</small><b>{person.status}</b></span><span><small>Coverage</small><b>{person.status === "Unmatched" ? "Unavailable" : person.syncEnabled ? "Enabled" : "Paused"}</b></span><span><small>Schoolbox ID</small><b>{String(detail.user.schoolboxUserId ?? "Unmatched")}</b></span><span><small>Managed data</small><b>{detail.events.length} event(s), {detail.calendars.length} calendar(s)</b></span></div>
      <div className="person-detail-tabs"><button className={section === "overview" ? "active" : ""} onClick={() => setSection("overview")}>Overview and diagnostics</button>{person.status !== "Unmatched" && <button className={section === "exclusions" ? "active" : ""} onClick={() => setSection("exclusions")}>Event exclusions{detail.exclusions.categories.length + detail.exclusions.eventTypes.length > 0 ? ` (${detail.exclusions.categories.length + detail.exclusions.eventTypes.length})` : ""}</button>}</div>
      {section === "overview" ? <>
        <h3>Identities</h3><dl className="diagnostic-grid"><div><dt>{TARGET_LABELS[person.target]}</dt><dd>{person.targetEmail}</dd></div><div><dt>Schoolbox</dt><dd>{person.schoolboxEmail}</dd></div><div><dt>{person.target === "google" ? "Google user ID" : "Microsoft Entra user ID"}</dt><dd><code>{person.id}</code></dd></div><div><dt>Role</dt><dd>{person.role}</dd></div></dl>
        {typeof detail.user.lastError === "string" && detail.user.lastError && <div className="diagnostic-error-box"><b>Latest user error</b><p>{detail.user.lastError}</p></div>}
        <h3>Relay-created calendars</h3>
        {detail.calendars.length ? <div className="diagnostic-calendars">{detail.calendars.map(calendar => <details key={calendar.destinationId}><summary><b>{calendar.summary}</b><span>{calendar.destinationId}</span></summary><dl className="diagnostic-grid"><div><dt>{person.target === "google" ? "Google calendar ID" : "Outlook calendar ID"}</dt><dd><code>{calendar.targetCalendarId || calendar.googleCalendarId}</code></dd></div><div><dt>Time zone</dt><dd>{calendar.timeZone}</dd></div><div><dt>Description</dt><dd>{calendar.description || "None"}</dd></div><div><dt>Updated</dt><dd>{diagnosticDate(calendar.updatedAt)}</dd></div></dl></details>)}</div> : <div className="diagnostic-empty">No Relay-created secondary calendars are recorded for {TARGET_LABELS[person.target]}.</div>}
        <h3>Current managed events</h3><DiagnosticEventList events={detail.events} empty="No managed events are currently recorded for this person." />
        <h3>Recent run outcomes</h3>
        {detail.runs.length ? <div className="diagnostic-run-list">{detail.runs.map(outcome => <div className={outcome.status === "failed" ? "failed" : ""} key={outcome.runId}><b>{outcome.status} · {diagnosticDate(outcome.completedAt || outcome.startedAt)}</b><small>{outcome.eventsFound} found · {outcome.eventsCreated} created · {outcome.eventsUpdated} updated · {outcome.eventsDeleted} deleted · {outcome.eventsUnchanged} unchanged</small>{outcome.errorMessage && <p>{outcome.errorMessage}</p>}</div>)}</div> : <div className="diagnostic-empty">Detailed per-run history starts with the next run on the enhanced diagnostic schema.</div>}
      </> : <UserExclusionEditor person={person} detail={detail} canConfigure={canConfigure} onSaved={onPreferencesSaved} />}
    </>}
  </aside></div>;
}

function RunsPage({ runs, selectedRun, setSelectedRun, runNow, syncRunning, canOperate, loadError }: { runs: Run[]; selectedRun: Run | null; setSelectedRun: (run: Run | null) => void; runNow: (targets?: TargetProvider[]) => Promise<void>; syncRunning: boolean; canOperate: boolean; loadError: boolean }) {
  const [status, setStatus] = useState("All statuses");
  const [detailUsers, setDetailUsers] = useState<RunUserDiagnostic[]>([]);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailError, setDetailError] = useState("");
  const [userQuery, setUserQuery] = useState("");
  const [selectedOutcome, setSelectedOutcome] = useState<RunUserDiagnostic | null>(null);
  const [events, setEvents] = useState<DiagnosticEvent[]>([]);
  const [eventTotal, setEventTotal] = useState(0);
  const [eventsLoading, setEventsLoading] = useState(false);
  const [targetFilter, setTargetFilter] = useState<"all" | TargetProvider>("all");
  const [detailTargets, setDetailTargets] = useState<RunTargetDiagnostic[]>([]);
  const filtered = runs.filter(run => status === "All statuses" || run.status === status);
  const failures = detailUsers.filter(user => user.status === "failed" && (targetFilter === "all" || (user.target ?? "google") === targetFilter));
  const visibleOutcomes = detailUsers.filter(user => (targetFilter === "all" || (user.target ?? "google") === targetFilter) && `${user.displayName ?? ""} ${user.targetEmail ?? user.googleEmail} ${user.schoolboxEmail ?? ""} ${user.stage} ${user.errorMessage ?? ""}`.toLowerCase().includes(userQuery.toLowerCase()));

  const loadOutcome = useCallback(async (outcome: RunUserDiagnostic) => {
    if (!selectedRun) return;
    setSelectedOutcome(outcome);
    setEvents([]);
    setEventTotal(0);
    setEventsLoading(true);
    try {
      const outcomeTarget = outcome.target ?? "google";
      const outcomeUserId = outcome.targetUserId ?? outcome.googleUserId;
      const payload = await fetchJson(`/api/run-details?runId=${encodeURIComponent(selectedRun.id)}&target=${outcomeTarget}&userId=${encodeURIComponent(outcomeUserId)}&limit=250`);
      setEvents((payload.events as DiagnosticEvent[] | undefined) ?? []);
      setEventTotal(Number(payload.total ?? 0));
    } catch (error) {
      setDetailError(error instanceof Error ? error.message : "Event diagnostics could not be loaded.");
    } finally {
      setEventsLoading(false);
    }
  }, [selectedRun]);

  const loadMoreEvents = async () => {
    if (!selectedRun || !selectedOutcome || eventsLoading || events.length >= eventTotal) return;
    setEventsLoading(true);
    try {
      const outcomeTarget = selectedOutcome.target ?? "google";
      const outcomeUserId = selectedOutcome.targetUserId ?? selectedOutcome.googleUserId;
      const payload = await fetchJson(`/api/run-details?runId=${encodeURIComponent(selectedRun.id)}&target=${outcomeTarget}&userId=${encodeURIComponent(outcomeUserId)}&limit=250&offset=${events.length}`);
      setEvents(current => [...current, ...((payload.events as DiagnosticEvent[] | undefined) ?? [])]);
      setEventTotal(Number(payload.total ?? eventTotal));
    } catch (error) {
      setDetailError(error instanceof Error ? error.message : "Additional event diagnostics could not be loaded.");
    } finally {
      setEventsLoading(false);
    }
  };

  useEffect(() => {
    let cancelled = false;
    if (!selectedRun) return () => { cancelled = true; };
    void (async () => {
      await Promise.resolve();
      if (cancelled) return;
      setDetailUsers([]);
      setSelectedOutcome(null);
      setDetailTargets([]);
      setEvents([]);
      setDetailError("");
      setUserQuery("");
      setTargetFilter("all");
      setDetailLoading(true);
      try {
        const payload = await fetchJson(`/api/run-details?runId=${encodeURIComponent(selectedRun.id)}`);
        if (cancelled) return;
        const outcomes = (payload.users as RunUserDiagnostic[] | undefined) ?? [];
        setDetailUsers(outcomes);
        setDetailTargets((payload.targets as RunTargetDiagnostic[] | undefined) ?? selectedRun.targets ?? []);
        const firstFailure = outcomes.find(outcome => outcome.status === "failed");
        if (firstFailure) void loadOutcome(firstFailure);
      } catch (error) {
        if (!cancelled) setDetailError(error instanceof Error ? error.message : "Run diagnostics could not be loaded.");
      } finally {
        if (!cancelled) setDetailLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [selectedRun, loadOutcome]);

  const downloadRun = (run: Run) => {
    const diagnostic = { run, users: detailUsers, selectedUser: selectedOutcome, events, eventTotal };
    const url = URL.createObjectURL(new Blob([JSON.stringify(diagnostic, null, 2)], { type: "application/json" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = `${run.id}-diagnostic.json`;
    link.click();
    URL.revokeObjectURL(url);
  };
  return <div className={`runs-layout ${selectedRun ? "with-drawer" : ""}`}>
    <div>
      <section className="panel runs-panel"><div className="people-tools"><div><h2>Run history</h2><p>All scheduled and manual sync attempts</p></div><select value={status} onChange={e => setStatus(e.target.value)} aria-label="Filter run status"><option>All statuses</option><option>Succeeded</option><option>Warning</option><option>Failed</option></select>{canOperate && <button className="button secondary" onClick={() => void runNow()} disabled={syncRunning}>{syncRunning ? "Starting…" : "Run diagnostic sync"}</button>}</div><div className="table-wrap"><table><thead><tr><th>Run</th><th>Started</th><th>Status</th><th>Synced</th><th>Changes</th><th>Duration</th><th><span className="sr-only">Open</span></th></tr></thead><tbody>{filtered.map(run => <RunRow key={run.id} run={run} onClick={() => setSelectedRun(run)} />)}{filtered.length === 0 && <tr><td colSpan={7} className="table-empty">{loadError ? "Run history could not be loaded." : "No matching runs. Start a sync when setup is complete."}</td></tr>}</tbody></table></div></section>
    </div>
    {selectedRun && <aside className="run-drawer enhanced-run-drawer">
      <div className="drawer-head"><div><p className="eyebrow">Run detail</p><h2>{selectedRun.id}</h2></div><button onClick={() => setSelectedRun(null)} aria-label="Close run details">×</button></div>
      <StatusPill status={selectedRun.status} />
      <div className="drawer-summary"><span><small>Started</small><b>{selectedRun.started}</b></span><span><small>Duration</small><b>{selectedRun.duration}</b></span><span><small>Trigger</small><b>{selectedRun.trigger}</b></span><span><small>People synced</small><b>{selectedRun.users}</b></span><span><small>Current phase</small><b>{selectedRun.phase.replaceAll("_", " ")}</b></span><span><small>Last progress</small><b>{selectedRun.progressAt ? new Date(selectedRun.progressAt).toLocaleTimeString("en-AU") : "Not recorded"}</b></span></div>
      <h3>{selectedRun.status === "Running" ? "Current operation" : "Run narrative"}</h3><p>{selectedRun.status === "Running" ? selectedRun.phaseDetail : selectedRun.note}</p>
      <div className="event-stats"><div><span className="stat-mark green">+</span><b>{selectedRun.created ?? 0}</b><small>Created</small></div><div><span className="stat-mark navy">↻</span><b>{selectedRun.updated ?? 0}</b><small>Updated</small></div><div><span className="stat-mark red">−</span><b>{selectedRun.deleted ?? 0}</b><small>Removed</small></div></div>
      {detailTargets.length > 0 && <><h3>Target outcomes</h3><div className="run-target-cards">{detailTargets.map(target => <button type="button" key={target.target} className={`${targetFilter === target.target ? "selected" : ""} ${target.status}`} onClick={() => { setTargetFilter(target.target); setSelectedOutcome(null); }}><span className={`provider-mark ${target.target}`}>{target.target === "google" ? "G" : "M"}</span><span><b>{TARGET_LABELS[target.target]}</b><small>{target.usersSynced} synced · {target.usersSelected} enabled · {target.errors} error(s)</small></span><span><b>{target.status.replaceAll("_", " ")}</b><small>{target.phase.replaceAll("_", " ")}</small></span></button>)}</div><button type="button" className="text-button" onClick={() => { setTargetFilter("all"); setSelectedOutcome(null); }}>Show all target outcomes</button></>}
      {detailLoading && <div className="diagnostic-empty">Loading per-user outcomes…</div>}
      {detailError && <div className="diagnostic-error-box"><b>Diagnostic detail</b><p>{detailError}</p></div>}
      {failures.length > 0 && <><h3>Accounts requiring attention</h3><div className="run-failures">{failures.map(failure => { const provider = failure.target ?? "google"; const id = failure.targetUserId ?? failure.googleUserId; const email = failure.targetEmail ?? failure.googleEmail; return <button key={`${provider}:${id}`} className={selectedOutcome && (selectedOutcome.target ?? "google") === provider && (selectedOutcome.targetUserId ?? selectedOutcome.googleUserId) === id ? "selected" : ""} onClick={() => void loadOutcome(failure)}><b>{failure.displayName || email}</b><small>{TARGET_LABELS[provider]} · {email} · {failure.stage.replaceAll("_", " ")}</small><p>{failure.errorMessage}</p></button>; })}</div></>}
      <h3>User outcomes</h3>
      {detailUsers.length > 0 ? <><div className="diagnostic-user-search"><input value={userQuery} onChange={event => setUserQuery(event.target.value)} placeholder="Search name, email, stage or error…" aria-label="Search run user outcomes" /><span>{visibleOutcomes.length} of {detailUsers.length}</span></div><div className="run-user-outcomes">{visibleOutcomes.map(outcome => { const provider = outcome.target ?? "google"; const id = outcome.targetUserId ?? outcome.googleUserId; const email = outcome.targetEmail ?? outcome.googleEmail; const selected = selectedOutcome && (selectedOutcome.target ?? "google") === provider && (selectedOutcome.targetUserId ?? selectedOutcome.googleUserId) === id; return <button key={`${provider}:${id}`} className={`${outcome.status} ${selected ? "selected" : ""}`} onClick={() => void loadOutcome(outcome)}><span><b>{outcome.displayName || email}</b><small>{TARGET_LABELS[provider]} · {email}</small></span><span><b>{outcome.status}</b><small>{outcome.eventsCreated} + · {outcome.eventsUpdated} ↻ · {outcome.eventsDeleted} −</small></span></button>; })}</div></> : !detailLoading && <div className="diagnostic-empty">Detailed user outcomes were not recorded for this older run.</div>}
      {selectedOutcome && <section className="selected-outcome"><h3>Selected account</h3><div className="selected-outcome-head"><div><b>{selectedOutcome.displayName || selectedOutcome.targetEmail || selectedOutcome.googleEmail}</b><small>{TARGET_LABELS[selectedOutcome.target ?? "google"]} · {selectedOutcome.targetEmail ?? selectedOutcome.googleEmail} · Schoolbox {selectedOutcome.schoolboxEmail || selectedOutcome.schoolboxUserId || "not matched"}</small></div><span className={selectedOutcome.status}>{selectedOutcome.status}</span></div>{selectedOutcome.errorMessage && <div className="diagnostic-error-box"><b>Failure during {selectedOutcome.stage.replaceAll("_", " ")}</b><p>{selectedOutcome.errorMessage}</p></div>}<dl className="diagnostic-grid"><div><dt>Events returned</dt><dd>{selectedOutcome.eventsFound}</dd></div><div><dt>Included by policy</dt><dd>{selectedOutcome.eventsIncluded}</dd></div><div><dt>Managed afterward</dt><dd>{selectedOutcome.managedEventsAfter}</dd></div><div><dt>Completed</dt><dd>{diagnosticDate(selectedOutcome.completedAt)}</dd></div></dl><h3>Event actions</h3>{eventsLoading && events.length === 0 ? <div className="diagnostic-empty">Loading event actions…</div> : <DiagnosticEventList events={events} empty="No event action was recorded before this user outcome." />}{eventTotal > events.length && <button className="button secondary full diagnostic-load-more" disabled={eventsLoading} onClick={() => void loadMoreEvents()}>{eventsLoading ? "Loading…" : `Load more events (${events.length} of ${eventTotal})`}</button>}</section>}
      {canOperate && selectedRun.status !== "Succeeded" && selectedRun.status !== "Running" && <button className="button primary full" onClick={() => void runNow()}>Retry enabled-user sync <span>→</span></button>}
      <button className="button ghost full" onClick={() => downloadRun(selectedRun)}>Download loaded diagnostic data</button>
    </aside>}
  </div>;
}

function SettingsPage({ initialSection, config, setConfig, saveConfig, setNotice, onOpenSetup }: {
  initialSection?: string;
  config: Config;
  setConfig: React.Dispatch<React.SetStateAction<Config>>;
  saveConfig: (message?: string) => Promise<boolean>;
  setNotice: (notice: Notice) => void;
  onOpenSetup: (target: Exclude<SetupTrack, "hub">) => void;
}) {
  const applicationOrigin = useApplicationOrigin();
  const microsoftConsentCallback = `${applicationOrigin || "https://relay-host"}/api/auth/microsoft/admin-consent/callback`;
  const sections = ["Schedule", "People", "Event rules", "Event content", "Connections", "Reconciliation", "Advanced"];
  const [section, setSection] = useState(initialSection ?? "Schedule");
  const [testing, setTesting] = useState<"schoolbox" | "google" | "microsoft" | null>(null);
  const [startingMicrosoftConsent, setStartingMicrosoftConsent] = useState(false);
  const [policyTarget, setPolicyTarget] = useState<TargetProvider>(config.googleEnabled ? "google" : "microsoft");
  const [eventTypes, setEventTypes] = useState<DiscoveredEventType[]>([]);
  const [calendarUsage, setCalendarUsage] = useState<CalendarDestinationUsage[]>([]);
  const [retiringCalendar, setRetiringCalendar] = useState<string | null>(null);
  const [typeRuleText, setTypeRuleText] = useState(() => (config.googleEnabled ? config.syncPolicy : config.microsoftSyncPolicy).eventTypes.join("\n"));
  const [microsoftCredentialsEdited, setMicrosoftCredentialsEdited] = useState(false);
  const [storedConnectionIdentity, setStoredConnectionIdentity] = useState({
    schoolboxUrl: config.schoolboxUrl,
    googleAdminEmail: config.adminEmail,
    googleCustomer: config.googleCustomer,
    microsoftTenantId: config.microsoftTenantId,
    microsoftClientId: config.microsoftClientId,
    microsoftTestUserEmail: config.microsoftTestUserEmail,
  });
  const policy = policyTarget === "google" ? config.syncPolicy : config.microsoftSyncPolicy;
  const schoolboxConnectionEdited = Boolean(config.schoolboxJwt) || config.schoolboxUrl !== storedConnectionIdentity.schoolboxUrl;
  const googleConnectionEdited = Boolean(config.serviceAccountJson) || config.adminEmail !== storedConnectionIdentity.googleAdminEmail || config.googleCustomer !== storedConnectionIdentity.googleCustomer;
  const microsoftConnectionEdited = microsoftCredentialsEdited || Boolean(config.microsoftClientSecret) || config.microsoftTenantId !== storedConnectionIdentity.microsoftTenantId || config.microsoftClientId !== storedConnectionIdentity.microsoftClientId || config.microsoftTestUserEmail !== storedConnectionIdentity.microsoftTestUserEmail;

  useEffect(() => {
    let cancelled = false;
    void fetchJson("/api/event-types").then((payload) => {
      if (!cancelled) setEventTypes((payload.eventTypes as DiscoveredEventType[] | undefined) ?? []);
    }).catch(() => undefined);
    void fetchJson(`/api/calendar-destinations?target=${policyTarget}`).then((payload) => {
      if (!cancelled) setCalendarUsage((payload.destinations as CalendarDestinationUsage[] | undefined) ?? []);
    }).catch(() => undefined);
    return () => { cancelled = true; };
  }, [policyTarget]);

  const changePolicyTarget = (target: TargetProvider) => {
    setPolicyTarget(target);
    setTypeRuleText((target === "google" ? config.syncPolicy : config.microsoftSyncPolicy).eventTypes.join("\n"));
  };
  useEffect(() => {
    const currentTypes = typeRuleText.split(/\r?\n/).map(item => item.trim()).filter(Boolean);
    if (currentTypes.length === policy.eventTypes.length && currentTypes.every((item, index) => item === policy.eventTypes[index])) return;
    const timer = window.setTimeout(() => setTypeRuleText(policy.eventTypes.join("\n")), 0);
    return () => window.clearTimeout(timer);
  }, [policy.eventTypes, typeRuleText]);

  const setPolicy = (update: Partial<SyncPolicy>) => setConfig(current => {
    const currentPolicy = policyTarget === "google" ? current.syncPolicy : current.microsoftSyncPolicy;
    const nextPolicy = normalizeSyncPolicy({ ...currentPolicy, ...update }, currentPolicy);
    return policyTarget === "google" ? { ...current, syncPolicy: nextPolicy } : { ...current, microsoftSyncPolicy: nextPolicy };
  });
  const setCategory = (category: EventCategory, enabled: boolean) => setConfig(current => {
    const currentPolicy = policyTarget === "google" ? current.syncPolicy : current.microsoftSyncPolicy;
    const nextPolicy = normalizeSyncPolicy({ ...currentPolicy, categories: { ...currentPolicy.categories, [category]: enabled } }, currentPolicy);
    return policyTarget === "google" ? { ...current, syncPolicy: nextPolicy } : { ...current, microsoftSyncPolicy: nextPolicy };
  });
  const updateTypeRules = (value: string) => {
    setTypeRuleText(value);
    setPolicy({ eventTypes: value.split(/\r?\n/).map(item => item.trim()).filter(Boolean) });
  };
  const setCategoryOverride = (category: EventCategory, rule: GoogleEventRuleOverride) => {
    setConfig(current => {
      const currentPolicy = policyTarget === "google" ? current.syncPolicy : current.microsoftSyncPolicy;
      const next = { ...currentPolicy.categoryOverrides };
      if (Object.keys(rule).length) next[category] = rule; else delete next[category];
      const nextPolicy = normalizeSyncPolicy({ ...currentPolicy, categoryOverrides: next }, currentPolicy);
      return policyTarget === "google" ? { ...current, syncPolicy: nextPolicy } : { ...current, microsoftSyncPolicy: nextPolicy };
    });
  };
  const setTypeOverride = (key: string, rule: GoogleEventRuleOverride) => {
    setConfig(current => {
      const currentPolicy = policyTarget === "google" ? current.syncPolicy : current.microsoftSyncPolicy;
      const next = { ...currentPolicy.eventTypeOverrides };
      if (Object.keys(rule).length) next[key] = rule; else delete next[key];
      const nextPolicy = normalizeSyncPolicy({ ...currentPolicy, eventTypeOverrides: next }, currentPolicy);
      return policyTarget === "google" ? { ...current, syncPolicy: nextPolicy } : { ...current, microsoftSyncPolicy: nextPolicy };
    });
  };
  const addCalendar = () => {
    const id = `calendar-${crypto.randomUUID().replaceAll("-", "").slice(0, 12)}`;
    setPolicy({ secondaryCalendars: [...policy.secondaryCalendars, {
      id,
      name: "Separate calendar",
      description: "Events synchronized from Schoolbox by Relay.",
    }] });
  };
  const updateCalendar = (id: string, update: Partial<ManagedCalendarDefinition>) => {
    // Keep the editor mounted while a field is temporarily blank during typing.
    setConfig(current => {
      const currentPolicy = policyTarget === "google" ? current.syncPolicy : current.microsoftSyncPolicy;
      const nextPolicy = { ...currentPolicy, secondaryCalendars: currentPolicy.secondaryCalendars.map(calendar => calendar.id === id ? { ...calendar, ...update } : calendar) };
      return policyTarget === "google" ? { ...current, syncPolicy: nextPolicy } : { ...current, microsoftSyncPolicy: nextPolicy };
    });
  };
  const removeCalendar = (id: string) => {
    const definition = policy.secondaryCalendars.find(calendar => calendar.id === id);
    if (!definition || !window.confirm(`Remove the destination “${definition.name}” from ${TARGET_LABELS[policyTarget]} routing? Existing provider calendars will remain and appear below for separate cleanup.`)) return;
    setConfig(current => policyTarget === "google"
      ? { ...current, syncPolicy: withoutManagedCalendarDestination(current.syncPolicy, id) }
      : { ...current, microsoftSyncPolicy: withoutManagedCalendarDestination(current.microsoftSyncPolicy, id) });
  };
  const retireCalendar = async (id: string, name: string, calendarCount: number) => {
    if (retiringCalendar) return;
    const confirmed = window.confirm(
      `Retire “${name}” and permanently delete ${calendarCount} tracked ${TARGET_CALENDAR_LABELS[policyTarget]} secondary calendar(s)? All content in those calendars, including manually added events, will be deleted. Relay will also remove this destination from this target's saved routing. This cannot be undone.`,
    );
    if (!confirmed) return;
    setRetiringCalendar(id);
    try {
      const payload = await fetchJson("/api/calendar-destinations", {
        method: "DELETE",
        body: JSON.stringify({ target: policyTarget, destinationId: id }),
      });
      setConfig(current => policyTarget === "google"
        ? { ...current, syncPolicy: withoutManagedCalendarDestination(current.syncPolicy, id) }
        : { ...current, microsoftSyncPolicy: withoutManagedCalendarDestination(current.microsoftSyncPolicy, id) });
      setCalendarUsage((payload.destinations as CalendarDestinationUsage[] | undefined) ?? []);
      const deleted = Number(payload.calendarsDeleted ?? 0);
      const alreadyMissing = Number(payload.calendarsAlreadyMissing ?? 0);
      const remaining = Number(payload.calendarsRemaining ?? 0);
      const removedEvents = Number(payload.eventMappingsRemoved ?? 0);
      if (remaining > 0 || payload.error) {
        setNotice({ kind: "error", message: `Destination routing was retired and ${deleted} calendar(s) deleted, but ${remaining} tracked calendar(s) still need cleanup. Retry after checking ${TARGET_LABELS[policyTarget]} access.` });
      } else {
        const missingNote = alreadyMissing > 0 ? ` ${alreadyMissing} calendar(s) were already absent.` : "";
        setNotice({ kind: "success", message: `Destination retired and ${deleted} ${TARGET_CALENDAR_LABELS[policyTarget]} calendar(s) deleted. ${removedEvents} tracked event mapping(s) removed.${missingNote}` });
      }
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "The calendar destination could not be retired." });
    } finally {
      setRetiringCalendar(null);
    }
  };
  const testConnection = async (target: "schoolbox" | "google" | "microsoft") => {
    if (target === "schoolbox" && schoolboxConnectionEdited) {
      setNotice({ kind: "error", message: "Save or discard the Schoolbox credential change before testing the stored connection." });
      return;
    }
    if (target === "google" && googleConnectionEdited) {
      setNotice({ kind: "error", message: "Save or discard the Google credential change before testing the stored connection." });
      return;
    }
    if (target === "microsoft" && microsoftConnectionEdited) {
      setNotice({ kind: "error", message: "Save or discard the Microsoft credential change before testing the stored connection." });
      return;
    }
    if (target === "microsoft" && !config.microsoftTestUserEmail.trim()) {
      setNotice({ kind: "error", message: "Enter a test mailbox before testing Microsoft 365." });
      return;
    }
    setTesting(target);
    try {
      const payload = await fetchJson("/api/diagnostics", {
        method: "POST",
        body: JSON.stringify({ target }),
      });
      if (target === "microsoft") {
        setConfig(current => ({ ...current, microsoftConsentGrantedAt: String(payload.microsoftConsentGrantedAt ?? current.microsoftConsentGrantedAt), microsoftSetupCompleted: Boolean(payload.microsoftSetupCompleted ?? current.microsoftSetupCompleted) }));
      }
      setNotice({ kind: "success", message: String(payload.message ?? `${target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target]} connection verified.`) });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "The connection check failed." });
    } finally {
      setTesting(null);
    }
  };
  const startMicrosoftConsent = async () => {
    if (microsoftConnectionEdited) {
      setNotice({ kind: "error", message: "Save the Microsoft credential changes before granting admin consent." });
      return;
    }
    if (!config.microsoftTenantId || !config.microsoftClientId || !config.hasMicrosoftClientSecret || !config.microsoftTestUserEmail.trim()) {
      setNotice({ kind: "error", message: "Save the tenant ID, application ID, client secret and test mailbox before granting admin consent." });
      return;
    }
    setStartingMicrosoftConsent(true);
    try {
      const payload = await fetchJson("/api/auth/microsoft/admin-consent/start", { method: "POST", body: "{}" });
      const consentUrl = typeof payload.url === "string" ? payload.url : "";
      if (!consentUrl) throw new Error("The Microsoft admin-consent URL was not returned.");
      window.location.assign(consentUrl);
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Microsoft admin consent could not be started." });
      setStartingMicrosoftConsent(false);
    }
  };
  const submit = (event: FormEvent) => {
    event.preventDefault();
    const invalidTarget = ([config.syncPolicy, config.microsoftSyncPolicy] as SyncPolicy[]).find(candidate => {
      const names = candidate.secondaryCalendars.map(calendar => calendar.name.trim().toLocaleLowerCase("en-AU"));
      return names.some(name => !name) || new Set(names).size !== names.length;
    });
    if (invalidTarget?.secondaryCalendars.some(calendar => !calendar.name.trim())) {
      setNotice({ kind: "error", message: "Give every secondary calendar destination a name before saving." });
      return;
    }
    if (invalidTarget) {
      setNotice({ kind: "error", message: "Each secondary calendar destination needs a unique name." });
      return;
    }
    const schoolboxCredentialsChanged = schoolboxConnectionEdited;
    const googleCredentialsChanged = googleConnectionEdited;
    const microsoftCredentialsChanged = microsoftConnectionEdited;
    void (async () => {
      if (!await saveConfig()) return;
      setStoredConnectionIdentity({
        schoolboxUrl: config.schoolboxUrl,
        googleAdminEmail: config.adminEmail,
        googleCustomer: config.googleCustomer,
        microsoftTenantId: config.microsoftTenantId,
        microsoftClientId: config.microsoftClientId,
        microsoftTestUserEmail: config.microsoftTestUserEmail,
      });
      setMicrosoftCredentialsEdited(false);
      if (schoolboxCredentialsChanged || googleCredentialsChanged || microsoftCredentialsChanged) {
        setNotice({ kind: "info", message: "Connection changes were saved. Complete the affected setup guide to re-test and re-enable that connection." });
      }
    })();
  };
  const categoryCopy = EVENT_CATEGORY_COPY;
  const destinationName = (destinationId: string) => destinationId === "primary"
    ? "Primary calendar"
    : policy.secondaryCalendars.find(calendar => calendar.id === destinationId)?.name ?? "Unknown destination";
  const googleRuleSummary = (category: EventCategory, type: string | null) => {
    const resolved = resolveGoogleEventRule({ category, type }, policy);
    return `${destinationName(resolved.destinationId)} · ${resolved.transparency === "opaque" ? "Busy" : "Available"}`;
  };
  const typeCoverageSummary = (entry: DiscoveredEventType) => {
    const override = policy.eventTypeOverrides[entry.key];
    if (override?.enabled === true) return "Included by type";
    if (override?.enabled === false) return "Excluded by type";
    if (!policy.categories[entry.category]) return "Excluded by category";
    if (policy.eventTypeMode === "all") return "Included";
    const listed = policy.eventTypes.some(type => eventTypeKey(type) === entry.key);
    if (policy.eventTypeMode === "include") return listed ? "Included by filter" : "Excluded by filter";
    return listed ? "Excluded by filter" : "Included by filter";
  };
  const openConnectionSetup = (target: Exclude<SetupTrack, "hub">) => {
    const edited = target === "schoolbox" ? schoolboxConnectionEdited : target === "google" ? googleConnectionEdited : microsoftConnectionEdited;
    if (edited) {
      setNotice({ kind: "error", message: `Save or refresh to discard the ${target === "schoolbox" ? "Schoolbox" : TARGET_LABELS[target]} connection changes before opening its setup guide.` });
      return;
    }
    onOpenSetup(target);
  };
  const configuredDestinationIds = new Set(policy.secondaryCalendars.map(calendar => calendar.id));
  const usageForDestination = (id: string) => calendarUsage.find(usage => usage.destinationId === id);
  const retiredCalendarUsage = calendarUsage.filter(usage => !configuredDestinationIds.has(usage.destinationId));

  return <div className="settings-layout">
    <aside className="settings-nav">{sections.map(item => <button type="button" key={item} className={section === item ? "active" : ""} onClick={() => setSection(item)}>{item}<span>→</span></button>)}</aside>
    <form className="panel settings-card" onSubmit={submit}>
      {section === "Schedule" && <SettingsSection title="Sync schedule" intro="Choose how frequently Relay checks Schoolbox and the rolling calendar window it maintains.">
        <div className="form-grid"><Field label="Frequency"><select value={config.interval} onChange={e => setConfig(c => ({ ...c, interval: e.target.value }))}><option value="15">Every 15 minutes</option><option value="30">Every 30 minutes</option><option value="60">Every hour</option><option value="180">Every 3 hours</option><option value="360">Every 6 hours</option><option value="720">Every 12 hours</option><option value="1440">Daily</option></select></Field><Field label="Keep past events"><select value={config.pastDays} onChange={e => setConfig(c => ({ ...c, pastDays: e.target.value }))}><option value="0">From today</option><option value="7">7 days</option><option value="14">14 days</option><option value="30">30 days</option><option value="60">60 days</option><option value="90">90 days</option><option value="180">180 days</option><option value="365">1 year</option></select></Field><Field label="Sync ahead"><select value={config.futureDays} onChange={e => setConfig(c => ({ ...c, futureDays: e.target.value }))}><option value="30">30 days</option><option value="60">60 days</option><option value="90">90 days</option><option value="180">180 days</option><option value="365">1 year</option><option value="730">2 years</option></select></Field></div>
        <div className="settings-note"><span>◷</span><div><b>{Number(config.pastDays) + Number(config.futureDays)}-day rolling window</b><small>Schoolbox requests are automatically split into month-sized ranges.</small></div></div>
      </SettingsSection>}

      {section === "People" && <SettingsSection title="New-account coverage" intro="Choose independently what happens when Relay first discovers a matched account in each target directory.">
        <div className="target-default-grid settings-target-defaults">
          <div className={!config.googleEnabled ? "disabled" : ""}><span className="provider-mark google">G</span><div className="target-default-title"><b>Google Workspace</b><small>{config.googleEnabled ? "Configured target" : "Target is off; the preference can be prepared in advance."}</small></div><PolicyToggle checked={config.syncNewGoogleUsersByDefault} onChange={enabled => setConfig(c => ({ ...c, syncNewGoogleUsersByDefault: enabled, syncNewUsersByDefault: enabled }))} title="Enable newly matched Google accounts" detail={config.syncNewGoogleUsersByDefault ? "New matched accounts start syncing immediately." : "New matched accounts appear paused for review."} /></div>
          <div className={!config.microsoftEnabled ? "disabled" : ""}><span className="provider-mark microsoft">M</span><div className="target-default-title"><b>Microsoft 365</b><small>{config.microsoftEnabled ? "Configured target" : "Target is off; the preference can be prepared in advance."}</small></div><PolicyToggle checked={config.syncNewMicrosoftUsersByDefault} onChange={enabled => setConfig(c => ({ ...c, syncNewMicrosoftUsersByDefault: enabled }))} title="Enable newly matched Microsoft accounts" detail={config.syncNewMicrosoftUsersByDefault ? "New matched accounts start syncing immediately." : "New matched accounts appear paused for review."} /></div>
        </div>
        <div className="callout warm"><span>i</span><div><b>This changes future discoveries only</b><p>Existing user selections stay as they are. Use People to enable or pause individuals or selected groups.</p></div></div>
      </SettingsSection>}

      {section === "Event rules" && <SettingsSection title="Event rules" intro={`Configure Schoolbox coverage, destination and provider-native appearance for ${TARGET_LABELS[policyTarget]}.`}>
        <ProviderPolicyTabs target={policyTarget} setTarget={changePolicyTarget} config={config} />
        <div className="rule-order" aria-label="Event rule precedence">
          <div><span>1</span><b>Defaults</b><small>Applied to every included event.</small></div>
          <i aria-hidden="true">→</i>
          <div><span>2</span><b>Category</b><small>Overrides defaults for a source category.</small></div>
          <i aria-hidden="true">→</i>
          <div><span>3</span><b>Exact type</b><small>Final override for one Schoolbox type.</small></div>
        </div>
        <div className="callout"><span>i</span><div><b>The most specific setting wins</b><p>An exact-type value overrides its category; a category value overrides the default. “Inherit” means use the value from the preceding level. Timed, all-day, and completed-item switches remain global safeguards.</p></div></div>

        <h3 className="settings-subhead">Global coverage safeguards</h3>
        <div className="policy-grid three"><PolicyToggle checked={policy.includeTimedEvents} onChange={enabled => setPolicy({ includeTimedEvents: enabled })} title="Timed events" detail="Events with start and end times." /><PolicyToggle checked={policy.includeAllDayEvents} onChange={enabled => setPolicy({ includeAllDayEvents: enabled })} title="All-day events" detail="Events represented by dates rather than times." /><PolicyToggle checked={policy.includeCompletedEvents} onChange={enabled => setPolicy({ includeCompletedEvents: enabled })} title="Completed items" detail="Task-like items marked completed in Schoolbox." /></div>

        <h3 className="settings-subhead">Calendar destinations</h3>
        <div className="settings-note"><span>{policyTarget === "google" ? "G" : "M"}</span><div><b>{TARGET_LABELS[policyTarget]} primary calendar is always protected</b><small>Secondary calendars are created lazily and tracked separately for this provider. Removing or retiring one never targets the other provider.</small></div></div>
        <div className="calendar-definitions">{policy.secondaryCalendars.map(calendar => {
          const usage = usageForDestination(calendar.id);
          const calendarCount = usage?.calendarCount ?? 0;
          return <div className="calendar-definition" key={calendar.id}>
            <div className="calendar-definition-head"><div><b>{calendar.name || "Unnamed destination"}</b><small>Relay destination ID: {calendar.id}</small><small className="calendar-usage">{calendarCount} tracked user calendar(s) · {usage?.eventCount ?? 0} managed event(s)</small></div><div className="calendar-definition-actions"><button type="button" className="row-delete" onClick={() => removeCalendar(calendar.id)} disabled={retiringCalendar !== null}>Remove routing</button>{calendarCount > 0 && <button type="button" className="row-delete destructive" onClick={() => void retireCalendar(calendar.id, calendar.name || calendar.id, calendarCount)} disabled={retiringCalendar !== null}>{retiringCalendar === calendar.id ? "Retiring…" : "Retire & delete"}</button>}</div></div>
            <div className="form-grid two"><Field label="Calendar name"><input required maxLength={100} value={calendar.name} onChange={e => updateCalendar(calendar.id, { name: e.target.value })} placeholder="Choose a name users will recognise" /></Field><Field label="Description"><input maxLength={500} value={calendar.description} onChange={e => updateCalendar(calendar.id, { description: e.target.value })} placeholder="Optional description" /></Field></div>
          </div>;
        })}</div>
        <button type="button" className="button secondary add-destination" onClick={addCalendar} disabled={policy.secondaryCalendars.length >= 20}>+ Add secondary calendar destination</button>
        {retiredCalendarUsage.length > 0 && <div className="retired-calendar-list"><h4>Retired destinations awaiting cleanup</h4><p>These destinations are no longer routed but still have tracked {TARGET_CALENDAR_LABELS[policyTarget]} calendars. Retry deletion after resolving provider access.</p>{retiredCalendarUsage.map(usage => <div className="retired-calendar" key={usage.destinationId}><div><b>{usage.summary || usage.destinationId}</b><small>{usage.calendarCount} tracked user calendar(s) · {usage.eventCount} managed event(s)</small><small>Relay destination ID: {usage.destinationId}</small></div><button type="button" className="button danger" onClick={() => void retireCalendar(usage.destinationId, usage.summary || usage.destinationId, usage.calendarCount)} disabled={retiringCalendar !== null}>{retiringCalendar === usage.destinationId ? "Deleting…" : "Delete remaining calendars"}</button></div>)}</div>}

        <h3 className="settings-subhead">Default {policyTarget === "google" ? "Google Calendar" : "Outlook Calendar"} behaviour</h3>
        <p className="settings-section-copy">This is the fallback for every included event. Category and exact-type cards below show their effective destination and availability.</p>
        <div className="form-grid two">
          <Field label="Destination"><select value={policy.defaultDestinationId} onChange={e => setPolicy({ defaultDestinationId: e.target.value })}><option value="primary">Primary calendar</option>{policy.secondaryCalendars.map(calendar => <option key={calendar.id} value={calendar.id}>{calendar.name}</option>)}</select></Field>
          <Field label="Availability"><select value={policy.transparency} onChange={e => setPolicy({ transparency: e.target.value as SyncPolicy["transparency"] })}><option value="opaque">Busy</option><option value="transparent">{policyTarget === "microsoft" ? "Free" : "Available"}</option></select></Field>
          <Field label={policyTarget === "microsoft" ? "Sensitivity" : "Visibility"}><select value={policy.visibility} onChange={e => setPolicy({ visibility: e.target.value as SyncPolicy["visibility"] })}><option value="default">{policyTarget === "microsoft" ? "Normal" : "Calendar default"}</option><option value="private">Private</option>{policyTarget === "google" ? <option value="public">Public details</option> : policy.visibility === "public" ? <option value="public">Normal (migrated public setting)</option> : null}</select></Field>
          {policyTarget === "google" && <Field label="Event colour"><select value={policy.colorId} onChange={e => setPolicy({ colorId: e.target.value })}><option value="">Calendar default</option>{CALENDAR_COLOURS.map(([id, name]) => <option key={id} value={id}>{name}</option>)}</select></Field>}
          <Field label="Reminders"><select value={policy.reminderMode} onChange={e => setPolicy({ reminderMode: e.target.value as SyncPolicy["reminderMode"] })}><option value="calendar_default">{policyTarget === "microsoft" ? "Use Outlook defaults" : "Use calendar defaults"}</option><option value="none">No reminders</option><option value="custom">One custom reminder</option></select></Field>
          {policy.reminderMode === "custom" && <Field label="Custom reminder"><div className="inline-fields"><input type="number" min={0} max={40320} value={policy.reminderMinutes} onChange={e => setPolicy({ reminderMinutes: Number(e.target.value) })} />{policyTarget === "google" ? <select value={policy.reminderMethod} onChange={e => setPolicy({ reminderMethod: e.target.value as SyncPolicy["reminderMethod"] })}><option value="popup">Popup minutes before</option><option value="email">Email minutes before</option></select> : <span className="inline-field-label">minutes before</span>}</div></Field>}
        </div>

        <h3 className="settings-subhead">Category rules</h3>
        <p className="settings-section-copy">Turn a whole category on or off, then optionally override its {TARGET_LABELS[policyTarget]} behaviour. The summary shows the effective result.</p>
        <div className="rule-list">{EVENT_CATEGORIES.map(category => { const rule = policy.categoryOverrides[category] ?? {}; const included = policy.categories[category]; return <details className="rule-card" key={category}><summary><div><b>{categoryCopy[category][0]}</b><small>{Object.keys(rule).length ? "Custom category behaviour" : `Inherits ${TARGET_LABELS[policyTarget]} defaults`}</small></div><span className="rule-card-state"><b>{included ? "Included" : "Excluded"}</b><small>{googleRuleSummary(category, null)}</small></span></summary><div className="rule-card-body category-rule-body"><PolicyToggle checked={included} onChange={enabled => setCategory(category, enabled)} title="Include this category" detail={categoryCopy[category][1]} /><EventRuleEditor target={policyTarget} rule={rule} calendars={policy.secondaryCalendars} onChange={next => setCategoryOverride(category, next)} /></div></details>; })}</div>

        <h3 className="settings-subhead">Exact Schoolbox type rules</h3>
        <p className="settings-section-copy">For most installations, include every type here and use each detected type card to make exceptions. The manual list remains available for types not yet detected.</p>
        <div className="form-grid two"><Field label="Baseline type coverage"><select value={policy.eventTypeMode} onChange={e => setPolicy({ eventTypeMode: e.target.value as EventTypeFilterMode })}><option value="all">Use category coverage for every type</option><option value="include">Only include manually listed types</option><option value="exclude">Exclude manually listed types</option></select></Field><Field label="Manual type list" hint="One exact Schoolbox type label per line. Exact type cards below take precedence."><textarea rows={5} value={typeRuleText} onChange={e => updateTypeRules(e.target.value)} disabled={policy.eventTypeMode === "all"} placeholder={policy.eventTypeMode === "include" ? "Timetable\nExcursion" : "Private appointment"} /></Field></div>
        {eventTypes.length ? <div className="rule-list">{eventTypes.map(entry => { const rule = policy.eventTypeOverrides[entry.key] ?? {}; return <details className="rule-card" key={entry.key}><summary><div><b>{entry.label}</b><small>{categoryCopy[entry.category]?.[0] ?? "Other and custom"} · {Object.keys(rule).length ? "Custom exact rule" : "Inherits category"}</small></div><span className="rule-card-state"><b>{typeCoverageSummary(entry)}</b><small>{googleRuleSummary(entry.category, entry.label)}</small></span></summary><div className="rule-card-body"><EventRuleEditor target={policyTarget} rule={rule} calendars={policy.secondaryCalendars} onChange={next => setTypeOverride(entry.key, next)} allowCoverage /></div></details>; })}</div> : <div className="detected-types"><div><b>No type labels have been catalogued yet</b><small>Run a pilot sync for an enabled user. Relay records only the type labels needed for configuration; this screen does not show user or event details.</small></div></div>}
        {policyTarget === "google" ? <div className="callout warm"><span>!</span><div><b>Secondary calendars need one additional delegated scope</b><p>Add <code>{SECONDARY_CALENDAR_SCOPE}</code> to the service account’s domain-wide delegation before assigning one. This scope is limited to calendars created by the app.</p><button type="button" className="text-button" onClick={() => { void navigator.clipboard.writeText(SCOPES.join(",")); setNotice({ kind: "success", message: "All required Google scopes copied." }); }}>Copy complete scope list</button></div></div> : <div className="callout"><span>i</span><div><b>Outlook uses provider-native appearance</b><p>Availability maps to Busy or Free, source visibility maps to Normal or Private sensitivity, and custom reminders use Outlook’s reminder minutes. Google event colours do not apply.</p></div></div>}
      </SettingsSection>}

      {section === "Event content" && <SettingsSection title="Event content" intro={`Choose what Relay copies from Schoolbox into each managed ${TARGET_CALENDAR_LABELS[policyTarget]} event.`}>
        <ProviderPolicyTabs target={policyTarget} setTarget={changePolicyTarget} config={config} />
        <h3 className="settings-subhead">Copied fields</h3>
        <div className="policy-grid"><PolicyToggle checked={policy.includeDescription} onChange={enabled => setPolicy({ includeDescription: enabled })} title="Description" detail="Copy the plain-text Schoolbox event detail." /><PolicyToggle checked={policy.includeLocation} onChange={enabled => setPolicy({ includeLocation: enabled })} title="Location" detail="Copy room or location metadata." /><PolicyToggle checked={policy.includeSchoolboxLink} onChange={enabled => setPolicy({ includeSchoolboxLink: enabled })} title="Schoolbox link" detail="Add an Open in Schoolbox source link." /><PolicyToggle checked={policy.includeEventTypeInDescription} onChange={enabled => setPolicy({ includeEventTypeInDescription: enabled })} title="Type in description" detail="Append the Schoolbox type label." /><PolicyToggle checked={policy.includeAuthorInDescription} onChange={enabled => setPolicy({ includeAuthorInDescription: enabled })} title="Author in description" detail="Append the source author when supplied." /></div>
        <h3 className="settings-subhead">Title</h3>
        <Field label="Title prefix" hint="Up to 40 characters; leave blank for the original title."><input maxLength={40} value={policy.titlePrefix} onChange={e => setPolicy({ titlePrefix: e.target.value })} placeholder="[Schoolbox]" /></Field>
        <div className="settings-note"><span>i</span><div><b>{TARGET_LABELS[policyTarget]} appearance follows the event-rule hierarchy</b><small>Configure destination, availability, sensitivity or visibility, and reminders together under Event rules.</small></div></div>
      </SettingsSection>}

      {section === "Connections" && <SettingsSection title="Connected services" intro="Review or replace every connection value configured during setup. Stored secrets are never revealed.">
        <div className="connection-settings-block"><div className="connection-settings-head"><span className="connection-logo">S</span><div><b>Schoolbox</b><small>{config.hasSchoolboxToken ? "JWT stored securely" : "Token required"}</small></div><button type="button" className="text-button" onClick={() => openConnectionSetup("schoolbox")}>Setup guide</button><button type="button" className="button secondary" onClick={() => void testConnection("schoolbox")} disabled={testing !== null}>{testing === "schoolbox" ? "Testing…" : "Test Schoolbox"}</button></div><div className="form-grid two"><Field label="Schoolbox base URL"><input type="url" value={config.schoolboxUrl} onChange={e => setConfig(c => ({ ...c, schoolboxUrl: e.target.value }))} placeholder="https://schoolbox.example.edu" /></Field><Field label="Replace superuser JWT" hint={config.hasSchoolboxToken ? "Leave blank to retain the encrypted token." : "Required before activation."}><input type="password" autoComplete="off" value={config.schoolboxJwt} onChange={e => setConfig(c => ({ ...c, schoolboxJwt: e.target.value }))} placeholder={config.hasSchoolboxToken ? "Stored securely" : "Paste Schoolbox JWT"} /></Field></div></div>
        <div className="connection-settings-block"><div className="connection-settings-head"><span className="connection-logo google">G</span><div><b>Google Workspace</b><small>{config.serviceAccountEmail || (config.hasGoogleServiceAccount ? "Service account stored securely" : "Credentials required")}</small><span className="connection-state-row"><i className={config.hasGoogleServiceAccount ? "ready" : "pending"}>{config.hasGoogleServiceAccount ? "Configured" : "Needs credentials"}</i><i className={config.googleSetupCompleted ? "ready" : "pending"}>{config.googleSetupCompleted ? "Setup complete" : "Setup required"}</i><i className={config.googleEnabled ? "enabled" : "off"}>{config.googleEnabled ? "Enabled" : "Disabled"}</i></span></div><button type="button" className="text-button" onClick={() => openConnectionSetup("google")}>Setup guide</button><button type="button" className="button secondary" onClick={() => void testConnection("google")} disabled={testing !== null}>{testing === "google" ? "Testing…" : "Test Google"}</button></div><PolicyToggle checked={config.googleEnabled} onChange={enabled => { if (enabled && (googleConnectionEdited || !config.googleSetupCompleted)) { setNotice({ kind: "error", message: googleConnectionEdited ? "Save or discard the Google connection changes, then complete its setup guide before enabling delivery." : "Complete the independent Google Workspace setup guide before enabling delivery." }); return; } setConfig(c => ({ ...c, googleEnabled: enabled })); }} title="Google Workspace delivery enabled" detail="Directory discovery and calendar synchronization run independently from Microsoft 365." /><div className="form-grid two"><Field label="Delegated administrator"><input type="email" value={config.adminEmail} onChange={e => { setConfig(c => ({ ...c, adminEmail: e.target.value })); }} placeholder="calendar-admin@example.edu" /></Field><Field label="Directory customer"><input value={config.googleCustomer} onChange={e => { setConfig(c => ({ ...c, googleCustomer: e.target.value })); }} placeholder="my_customer" /></Field></div><Field label="Replace service-account JSON" hint={config.hasGoogleServiceAccount ? `Leave blank to retain the encrypted credential${config.serviceAccountClientId ? ` (client ID ${config.serviceAccountClientId})` : ""}.` : "Paste the complete downloaded JSON key."}><textarea rows={7} value={config.serviceAccountJson} onChange={e => { setConfig(c => ({ ...c, serviceAccountJson: e.target.value })); }} placeholder={config.hasGoogleServiceAccount ? "Stored securely" : '{\n  "type": "service_account"\n}'} /></Field></div>
        <div className="connection-settings-block microsoft-connection"><div className="connection-settings-head"><span className="connection-logo microsoft">M</span><div><b>Microsoft 365</b><small>{config.microsoftTenantId || "Microsoft Entra application not configured"}</small><span className="connection-state-row"><i className={config.hasMicrosoftClientSecret ? "ready" : "pending"}>{config.hasMicrosoftClientSecret ? "Configured" : "Needs credentials"}</i><i className={config.microsoftSetupCompleted ? "ready" : "pending"}>{config.microsoftSetupCompleted ? "Setup complete" : microsoftCredentialsEdited ? "Save changes" : "Consent and test required"}</i><i className={config.microsoftEnabled ? "enabled" : "off"}>{config.microsoftEnabled ? "Enabled" : "Disabled"}</i></span></div><button type="button" className="text-button" onClick={() => openConnectionSetup("microsoft")}>Setup guide</button><button type="button" className="button secondary" onClick={() => void testConnection("microsoft")} disabled={testing !== null || startingMicrosoftConsent || !config.microsoftTestUserEmail.trim()}>{testing === "microsoft" ? "Testing…" : "Test Microsoft"}</button></div>
          <PolicyToggle checked={config.microsoftEnabled} onChange={enabled => { if (enabled && (microsoftConnectionEdited || !config.microsoftSetupCompleted)) { setNotice({ kind: "error", message: microsoftConnectionEdited ? "Save or discard the Microsoft connection changes, then complete its setup guide before enabling delivery." : "Complete the independent Microsoft 365 setup guide before enabling delivery." }); return; } setConfig(c => ({ ...c, microsoftEnabled: enabled })); }} title="Microsoft 365 delivery enabled" detail="Keep this off while credentials and admin consent are being prepared. Google Workspace continues independently." />
          <div className="form-grid two"><Field label="Directory (tenant) ID"><input value={config.microsoftTenantId} onChange={event => { setMicrosoftCredentialsEdited(true); setConfig(current => ({ ...current, microsoftTenantId: event.target.value })); }} placeholder="00000000-0000-0000-0000-000000000000" /></Field><Field label="Application (client) ID"><input value={config.microsoftClientId} onChange={event => { setMicrosoftCredentialsEdited(true); setConfig(current => ({ ...current, microsoftClientId: event.target.value })); }} placeholder="00000000-0000-0000-0000-000000000000" /></Field></div>
          <div className="form-grid two"><Field label="Replace client secret" hint={config.hasMicrosoftClientSecret ? "Leave blank to retain the encrypted secret." : "Paste the secret value, not its secret ID."}><input type="password" autoComplete="off" value={config.microsoftClientSecret} onChange={event => { setMicrosoftCredentialsEdited(true); setConfig(current => ({ ...current, microsoftClientSecret: event.target.value })); }} placeholder={config.hasMicrosoftClientSecret ? "Stored securely" : "Microsoft Entra client secret value"} /></Field><Field label="Test mailbox" hint="Required. Used only to verify Outlook calendar access."><input required type="email" value={config.microsoftTestUserEmail} onChange={event => { setMicrosoftCredentialsEdited(true); setConfig(current => ({ ...current, microsoftTestUserEmail: event.target.value })); }} placeholder="relay-test@example.edu" /></Field></div>
          <Field label="Web redirect URI" hint="Register this exact URI under Authentication → Web in Microsoft Entra before using admin consent."><CopyBox value={microsoftConsentCallback} onCopy={() => void navigator.clipboard.writeText(microsoftConsentCallback)} /></Field>
          <div className="microsoft-permissions"><b>Required Microsoft Graph application permissions</b><span>User.Read.All</span><span>Calendars.ReadWrite</span></div>
          <div className="consent-actions">{microsoftConnectionEdited || !config.hasMicrosoftClientSecret ? <button type="submit" className="button ghost" disabled={!config.microsoftTenantId || !config.microsoftClientId || (!config.microsoftClientSecret && !config.hasMicrosoftClientSecret) || !config.microsoftTestUserEmail.trim()}>Save Microsoft credentials</button> : <button type="button" className="button ghost" onClick={() => void startMicrosoftConsent()} disabled={startingMicrosoftConsent || !config.microsoftTestUserEmail.trim()}>{startingMicrosoftConsent ? "Opening Microsoft…" : "Grant or renew admin consent"} <span>↗</span></button>}<small>{microsoftConnectionEdited ? "Save the credential changes first. Relay will pause Microsoft delivery before consent or verification continues." : config.microsoftConsentGrantedAt ? `Admin consent and Microsoft Graph access verified ${diagnosticDate(config.microsoftConsentGrantedAt)}. Re-test after changing credentials.` : "Save the credentials with Microsoft delivery disabled, then return here to grant tenant admin consent."}</small></div>
        </div>
        <Field label="Calendar time zone" hint="IANA time-zone name used for timed Google and Outlook calendar events."><input value={config.timezone} onChange={e => setConfig(c => ({ ...c, timezone: e.target.value }))} placeholder="Australia/Sydney" /></Field>
        <div className="callout warm"><span>!</span><div><b>Credential changes use a deliberate re-verification sequence</b><p>Changing the Schoolbox host requires a replacement JWT. Leaving a secret field blank preserves it. Saving a changed Microsoft tenant, client ID or client secret automatically disables Microsoft delivery and clears its verified state. Save the replacement, renew consent if required, run Test Microsoft, then re-enable Microsoft delivery and save again. Secret-only rotation normally retains existing Entra admin consent.</p></div></div>
      </SettingsSection>}

      {section === "Reconciliation" && <SettingsSection title="Reconciliation and removal" intro={`Decide what Relay does with ${TARGET_LABELS[policyTarget]} events it previously created when the Schoolbox source or policy changes.`}>
        <ProviderPolicyTabs target={policyTarget} setTarget={changePolicyTarget} config={config} />
        <PolicyToggle checked={policy.deleteMissingEvents} onChange={enabled => setPolicy({ deleteMissingEvents: enabled })} title="Remove events no longer returned by Schoolbox" detail="Recommended for a true mirror. Removal is limited to Relay mapping records and the fetched date window for this target." />
        <PolicyToggle checked={policy.deleteExcludedEvents} onChange={enabled => setPolicy({ deleteExcludedEvents: enabled })} title="Remove events excluded by these settings" detail={`When a category, exact type, time form, or completion state is turned off, remove its existing Relay-managed ${TARGET_CALENDAR_LABELS[policyTarget]} event on the next enabled-account sync.`} />
        <div className="callout warm"><span>!</span><div><b>Policy removals are target-specific calendar changes</b><p>Save first, review the settings, then run a pilot account. Relay never targets manually created or third-party events, and never crosses into the other provider.</p></div></div>
      </SettingsSection>}

      {section === "Advanced" && <SettingsSection title="Advanced operations" intro="Control scheduler state and how much parallel work Relay sends to the APIs.">
        <PolicyToggle checked={config.enabled} onChange={enabled => setConfig(c => ({ ...c, enabled }))} title="Scheduled synchronization enabled" detail={config.enabled ? "The local scheduler starts runs at the configured interval." : "Scheduled runs are paused; manual runs remain available to operators."} />
        <Field label="Concurrent target accounts" hint="Applied separately within Google Workspace and Microsoft 365 branches. Lower this if either provider begins throttling; range 1–10."><input type="number" min={1} max={10} value={config.concurrency} onChange={e => setConfig(c => ({ ...c, concurrency: e.target.value }))} /></Field>
        <h3 className="settings-subhead">Failure limits</h3>
        <div className="form-grid">
          <Field label="Initial discovery timeout" hint="Maximum time for Schoolbox and each enabled target directory discovery; range 30–900 seconds."><input type="number" min={30} max={900} value={config.discoveryTimeoutSeconds} onChange={e => setConfig(c => ({ ...c, discoveryTimeoutSeconds: e.target.value }))} /></Field>
          <Field label="Per-account sync timeout" hint="Maximum time for one enabled target account; range 30–1800 seconds."><input type="number" min={30} max={1800} value={config.userSyncTimeoutSeconds} onChange={e => setConfig(c => ({ ...c, userSyncTimeoutSeconds: e.target.value }))} /></Field>
          <Field label="Whole-run timeout" hint="Hard limit for an organization run; range 5–240 minutes."><input type="number" min={5} max={240} value={config.runTimeoutMinutes} onChange={e => setConfig(c => ({ ...c, runTimeoutMinutes: e.target.value }))} /></Field>
        </div>
        <div className="callout"><span>✓</span><div><b>Stalled calls are aborted</b><p>API retries use exponential backoff, initial discovery and each user have deadlines, and the whole-run limit prevents a live heartbeat from masking a job that is making no progress.</p></div></div>
      </SettingsSection>}

      <div className="settings-actions"><span>{microsoftCredentialsEdited ? "Microsoft credential changes save with that target disabled. Test it, re-enable it, then save again." : "Changes take effect on the next sync. Saving does not start a run."}</span><button className="button primary" type="submit">{microsoftCredentialsEdited ? "Save credential changes" : "Save all settings"}</button></div>
    </form>
  </div>;
}

function ProviderPolicyTabs({ target, setTarget, config }: { target: TargetProvider; setTarget: (target: TargetProvider) => void; config: Config }) {
  return <div className="provider-policy-tabs" role="tablist" aria-label="Event policy target">{(["google", "microsoft"] as TargetProvider[]).map(provider => {
    const enabled = provider === "google" ? config.googleEnabled : config.microsoftEnabled;
    return <button type="button" role="tab" aria-selected={target === provider} className={target === provider ? "active" : ""} onClick={() => setTarget(provider)} key={provider}><span className={`provider-mark ${provider}`}>{provider === "google" ? "G" : "M"}</span><span><b>{TARGET_LABELS[provider]}</b><small>{enabled ? "Enabled target" : "Configure rules before enabling"}</small></span></button>;
  })}</div>;
}

function EventRuleEditor({ target, rule, calendars, onChange, allowCoverage = false }: {
  target: TargetProvider;
  rule: GoogleEventRuleOverride;
  calendars: ManagedCalendarDefinition[];
  onChange: (rule: GoogleEventRuleOverride) => void;
  allowCoverage?: boolean;
}) {
  const update = <K extends keyof GoogleEventRuleOverride>(field: K, value: GoogleEventRuleOverride[K] | undefined) => {
    const next = { ...rule };
    if (value === undefined) delete next[field]; else next[field] = value;
    onChange(next);
  };
  const colourValue = rule.colorId === undefined ? "inherit" : rule.colorId === "" ? "calendar_default" : rule.colorId;
  return <div className="rule-editor">
    {allowCoverage && <Field label="Sync coverage"><select value={rule.enabled === undefined ? "inherit" : String(rule.enabled)} onChange={e => update("enabled", e.target.value === "inherit" ? undefined : e.target.value === "true")}><option value="inherit">Inherit source coverage</option><option value="true">Always include this type</option><option value="false">Exclude this type</option></select></Field>}
    <Field label="Destination"><select value={rule.destinationId ?? "inherit"} onChange={e => update("destinationId", e.target.value === "inherit" ? undefined : e.target.value)}><option value="inherit">Inherit destination</option><option value="primary">Primary calendar</option>{calendars.map(calendar => <option key={calendar.id} value={calendar.id}>{calendar.name}</option>)}</select></Field>
    <Field label="Availability"><select value={rule.transparency ?? "inherit"} onChange={e => update("transparency", e.target.value === "inherit" ? undefined : e.target.value as GoogleEventRuleOverride["transparency"])}><option value="inherit">Inherit availability</option><option value="opaque">Busy</option><option value="transparent">{target === "microsoft" ? "Free" : "Available"}</option></select></Field>
    <Field label={target === "microsoft" ? "Sensitivity" : "Visibility"}><select value={rule.visibility ?? "inherit"} onChange={e => update("visibility", e.target.value === "inherit" ? undefined : e.target.value as GoogleEventRuleOverride["visibility"])}><option value="inherit">Inherit {target === "microsoft" ? "sensitivity" : "visibility"}</option><option value="default">{target === "microsoft" ? "Normal" : "Calendar default"}</option><option value="private">Private</option>{target === "google" ? <option value="public">Public details</option> : rule.visibility === "public" ? <option value="public">Normal (migrated public setting)</option> : null}</select></Field>
    {target === "google" && <Field label="Event colour"><select value={colourValue} onChange={e => update("colorId", e.target.value === "inherit" ? undefined : e.target.value === "calendar_default" ? "" : e.target.value)}><option value="inherit">Inherit colour</option><option value="calendar_default">Calendar default</option>{CALENDAR_COLOURS.map(([id, name]) => <option key={id} value={id}>{name}</option>)}</select></Field>}
    <Field label="Reminders"><select value={rule.reminderMode ?? "inherit"} onChange={e => update("reminderMode", e.target.value === "inherit" ? undefined : e.target.value as GoogleEventRuleOverride["reminderMode"])}><option value="inherit">Inherit reminders</option><option value="calendar_default">{target === "microsoft" ? "Use Outlook defaults" : "Use calendar defaults"}</option><option value="none">No reminders</option><option value="custom">One custom reminder</option></select></Field>
    {rule.reminderMode === "custom" && <Field label="Custom reminder"><div className="inline-fields"><input type="number" min={0} max={40320} value={rule.reminderMinutes ?? 10} onChange={e => update("reminderMinutes", Number(e.target.value))} />{target === "google" ? <select value={rule.reminderMethod ?? "popup"} onChange={e => update("reminderMethod", e.target.value as GoogleEventRuleOverride["reminderMethod"])}><option value="popup">Popup minutes before</option><option value="email">Email minutes before</option></select> : <span className="inline-field-label">minutes before</span>}</div></Field>}
  </div>;
}

function PolicyToggle({ checked, onChange, title, detail }: { checked: boolean; onChange: (checked: boolean) => void; title: string; detail: string }) {
  return <label className="settings-toggle"><span><b>{title}</b><small>{detail}</small></span><input type="checkbox" checked={checked} onChange={event => onChange(event.target.checked)} /><i aria-hidden="true" /></label>;
}

function AccessPage({ canChangeLocalPassword, setNotice, onSignedOut }: { canChangeLocalPassword: boolean; setNotice: (notice: Notice) => void; onSignedOut: () => void }) {
  const [staff, setStaff] = useState<StaffAccount[]>([]);
  const [settings, setSettings] = useState<OAuthSettings | null>(null);
  const [oauthForm, setOauthForm] = useState({ clientId: "", clientSecret: "", workspaceDomain: "" });
  const [newStaff, setNewStaff] = useState({ email: "", displayName: "", role: "viewer" as StaffRole });
  const [passwords, setPasswords] = useState({ currentPassword: "", nextPassword: "", confirmation: "" });
  const [busy, setBusy] = useState("");
  const [loadState, setLoadState] = useState<"loading" | "ready" | "error">("loading");

  const loadAccess = useCallback(async () => {
    setLoadState("loading");
    try {
      const [staffPayload, settingsPayload] = await Promise.all([fetchJson("/api/admin/staff"), fetchJson("/api/admin/auth-settings")]);
      const loadedStaff = (staffPayload.staff as StaffAccount[] | undefined) ?? [];
      const loadedSettings = settingsPayload.settings as OAuthSettings;
      setStaff(loadedStaff);
      setSettings(loadedSettings);
      setOauthForm(current => ({ ...current, clientId: loadedSettings.clientId, workspaceDomain: loadedSettings.workspaceDomain }));
      setLoadState("ready");
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) {
        onSignedOut();
        return;
      }
      setLoadState("error");
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Access settings could not be loaded." });
    }
  }, [onSignedOut, setNotice]);

  useEffect(() => {
    const timer = window.setTimeout(() => void loadAccess(), 0);
    return () => window.clearTimeout(timer);
  }, [loadAccess]);

  const saveOAuth = async (event: FormEvent) => {
    event.preventDefault();
    setBusy("oauth");
    try {
      const payload = await fetchJson("/api/admin/auth-settings", { method: "PUT", body: JSON.stringify(oauthForm) });
      setSettings(payload.settings as OAuthSettings);
      setOauthForm(current => ({ ...current, clientSecret: "" }));
      setNotice({ kind: "success", message: "Google Workspace sign-in settings saved." });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "OAuth settings could not be saved." });
    } finally { setBusy(""); }
  };

  const addStaff = async (event: FormEvent) => {
    event.preventDefault();
    setBusy("new-staff");
    try {
      await fetchJson("/api/admin/staff", { method: "PUT", body: JSON.stringify({ ...newStaff, enabled: true }) });
      setNewStaff({ email: "", displayName: "", role: "viewer" });
      await loadAccess();
      setNotice({ kind: "success", message: "IT staff access added. They can now sign in with Google Workspace." });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Staff access could not be added." });
    } finally { setBusy(""); }
  };

  const saveStaff = async (account: StaffAccount) => {
    setBusy(account.id);
    try {
      await fetchJson("/api/admin/staff", { method: "PUT", body: JSON.stringify(account) });
      await loadAccess();
      setNotice({ kind: "success", message: `${account.email} access updated.` });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Staff access could not be updated." });
    } finally { setBusy(""); }
  };

  const removeStaff = async (account: StaffAccount) => {
    if (!window.confirm(`Remove Relay access for ${account.email}?`)) return;
    setBusy(account.id);
    try {
      await fetchJson(`/api/admin/staff?id=${encodeURIComponent(account.id)}`, { method: "DELETE" });
      await loadAccess();
      setNotice({ kind: "success", message: `${account.email} access removed.` });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Staff access could not be removed." });
    } finally { setBusy(""); }
  };

  const changePassword = async (event: FormEvent) => {
    event.preventDefault();
    if (passwords.nextPassword !== passwords.confirmation) {
      setNotice({ kind: "error", message: "The new password confirmation does not match." });
      return;
    }
    setBusy("password");
    try {
      await fetchJson("/api/admin/password", { method: "PUT", body: JSON.stringify(passwords) });
      activeCsrfToken = "";
      onSignedOut();
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "The password could not be changed." });
    } finally { setBusy(""); }
  };

  const updateStaff = (id: string, update: Partial<StaffAccount>) => setStaff(current => current.map(account => account.id === id ? { ...account, ...update } : account));

  if (loadState !== "ready") return <section className="panel access-panel"><div className="empty-state"><b>{loadState === "loading" ? "Loading IT access…" : "IT access could not be loaded"}</b><p>{loadState === "loading" ? "Reading encrypted sign-in settings and staff permissions." : "The forms are disabled until the server responds successfully."}</p>{loadState === "error" && <button className="button secondary" onClick={() => void loadAccess()}>Try again</button>}</div></section>;

  return <div className="access-stack">
    <section className="panel access-panel">
      <div className="settings-head"><p className="eyebrow">Administrators</p><h2>Google Workspace sign-in</h2><p>Create an Internal Web OAuth client in Google Cloud, then enter it here. This is separate from the service account used for calendar synchronization.</p></div>
      <div className="access-guide"><ol><li><span>1</span><div><b>Configure the OAuth consent screen</b><p>Use an Internal audience so only accounts in your Workspace can authenticate.</p></div></li><li><span>2</span><div><b>Create a Web application OAuth client</b><p>Add the exact callback URL below as an authorised redirect URI.</p></div></li><li><span>3</span><div><b>Add staff to the allowlist</b><p>Workspace membership alone never grants Relay access.</p></div></li></ol></div>
      <Field label="Authorised redirect URI"><CopyBox value={settings?.callbackUrl ?? "Loading…"} onCopy={() => navigator.clipboard.writeText(settings?.callbackUrl ?? "")} /></Field>
      <form className="oauth-settings-form" onSubmit={saveOAuth}>
        <div className="form-grid"><Field label="Workspace domain"><input value={oauthForm.workspaceDomain} onChange={event => setOauthForm(current => ({ ...current, workspaceDomain: event.target.value }))} placeholder="school.edu.au" required /></Field><Field label="OAuth client ID"><input value={oauthForm.clientId} onChange={event => setOauthForm(current => ({ ...current, clientId: event.target.value }))} placeholder="123….apps.googleusercontent.com" required /></Field></div>
        <Field label="OAuth client secret" hint={settings?.hasClientSecret ? "A secret is stored. Leave blank to keep it unchanged." : "Encrypted before it is stored."}><input type="password" autoComplete="off" value={oauthForm.clientSecret} onChange={event => setOauthForm(current => ({ ...current, clientSecret: event.target.value }))} placeholder={settings?.hasClientSecret ? "Stored securely" : "Paste the client secret"} /></Field>
        <div className="settings-actions"><span className={`status-pill ${settings?.configured ? "success" : "warning"}`}><i />{settings?.configured ? "Google sign-in configured" : "Not configured"}</span><button className="button primary" type="submit" disabled={busy === "oauth"}>{busy === "oauth" ? "Saving…" : "Save Google sign-in"}</button></div>
      </form>
    </section>

    <section className="panel access-panel">
      <div className="settings-head"><h2>IT staff access</h2><p>Pre-approve individual Workspace identities and assign the least privilege they need.</p></div>
      <div className="role-grid"><div><b>Viewer</b><small>Dashboard, people, and run history</small></div><div><b>Operator</b><small>Viewer access plus diagnostics and manual syncs</small></div><div><b>Administrator</b><small>Connections, sync settings, and IT staff access</small></div></div>
      <form className="staff-add-form" onSubmit={addStaff}><Field label="Google Workspace email"><input type="email" value={newStaff.email} onChange={event => setNewStaff(current => ({ ...current, email: event.target.value }))} placeholder="it.staff@school.edu.au" required /></Field><Field label="Display name"><input value={newStaff.displayName} onChange={event => setNewStaff(current => ({ ...current, displayName: event.target.value }))} placeholder="Optional" /></Field><Field label="Role"><select value={newStaff.role} onChange={event => setNewStaff(current => ({ ...current, role: event.target.value as StaffRole }))}><option value="viewer">Viewer</option><option value="operator">Operator</option><option value="admin">Administrator</option></select></Field><button className="button primary" type="submit" disabled={busy === "new-staff"}>{busy === "new-staff" ? "Adding…" : "Add staff"}</button></form>
      <div className="staff-list">{staff.map(account => <div className="staff-row" key={account.id}><span className="person-avatar">{(account.displayName || account.email).split(/\s|@/).map(part => part[0]).join("").slice(0, 2).toUpperCase()}</span><div className="staff-identity"><b>{account.displayName || account.email}</b><small>{account.email} · {account.linked ? `Linked · Last login ${account.lastLoginAt ? new Date(account.lastLoginAt).toLocaleString("en-AU") : "not recorded"}` : "Awaiting first Google sign-in"}</small></div><select value={account.role} onChange={event => updateStaff(account.id, { role: event.target.value as StaffRole })} aria-label={`Role for ${account.email}`}><option value="viewer">Viewer</option><option value="operator">Operator</option><option value="admin">Administrator</option></select><label className="enable-control"><input type="checkbox" checked={account.enabled} onChange={event => updateStaff(account.id, { enabled: event.target.checked })} />Enabled</label><button className="button secondary" onClick={() => void saveStaff(account)} disabled={busy === account.id}>Save</button><button className="row-delete" onClick={() => void removeStaff(account)} disabled={busy === account.id}>Remove</button></div>)}{staff.length === 0 && <div className="empty-state"><b>No Google Workspace staff added</b><p>Add an email address above. The account remains blocked until explicitly listed here.</p></div>}</div>
    </section>

    {canChangeLocalPassword && <section className="panel access-panel narrow-panel">
      <div className="settings-head"><p className="eyebrow">Break-glass account</p><h2>Local administrator password</h2><p>Changing this password signs out the current session. Keep the credential in your IT password vault.</p></div>
      <form className="password-form" onSubmit={changePassword}><Field label="Current password"><input type="password" autoComplete="current-password" value={passwords.currentPassword} onChange={event => setPasswords(current => ({ ...current, currentPassword: event.target.value }))} required /></Field><Field label="New password" hint="At least 14 characters."><input type="password" autoComplete="new-password" minLength={14} value={passwords.nextPassword} onChange={event => setPasswords(current => ({ ...current, nextPassword: event.target.value }))} required /></Field><Field label="Confirm new password"><input type="password" autoComplete="new-password" minLength={14} value={passwords.confirmation} onChange={event => setPasswords(current => ({ ...current, confirmation: event.target.value }))} required /></Field><button className="button secondary" type="submit" disabled={busy === "password"}>{busy === "password" ? "Changing…" : "Change password"}</button></form>
    </section>}
  </div>;
}

function SettingsSection({ title, intro, children }: { title: string; intro: string; children: React.ReactNode }) { return <><div className="settings-head"><h2>{title}</h2><p>{intro}</p></div><div className="settings-content">{children}</div></>; }
