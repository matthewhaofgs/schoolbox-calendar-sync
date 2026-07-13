"use client";

import { FormEvent, useCallback, useEffect, useMemo, useRef, useState } from "react";

type View = "dashboard" | "setup" | "people" | "runs" | "settings" | "access";
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
  id: string;
  name: string;
  schoolboxEmail: string;
  googleEmail: string;
  role: string;
  status: "Synced" | "Syncing" | "Pending" | "Unmatched" | "Error";
  syncEnabled: boolean;
  lastSync: string;
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
  created?: number;
  updated?: number;
  deleted?: number;
  errors?: number;
};

const SCOPES = [
  "https://www.googleapis.com/auth/calendar.events.owned",
  "https://www.googleapis.com/auth/admin.directory.user.readonly",
];

function normalisePeople(value: unknown): Person[] | null {
  if (!Array.isArray(value)) return null;
  return value.map((item, index) => {
    const row = item as Record<string, unknown>;
    const statusValue = String(row.status ?? "Syncing").toLowerCase();
    const status: Person["status"] = statusValue === "synced" ? "Synced" : statusValue === "pending" ? "Pending" : statusValue === "unmatched" ? "Unmatched" : statusValue === "error" || statusValue === "failed" ? "Error" : "Syncing";
    return {
      id: String(row.id ?? row.googleUserId ?? `user-${index}`),
      name: String(row.name ?? row.displayName ?? "Unknown user"),
      schoolboxEmail: String(row.schoolboxEmail ?? row.sourceEmail ?? row.email ?? "—"),
      googleEmail: String(row.googleEmail ?? row.targetEmail ?? row.email ?? "—"),
      role: String(row.role ?? "User"),
      status,
      syncEnabled: row.syncEnabled === undefined ? true : Boolean(row.syncEnabled),
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
      note: String(row.note ?? row.message ?? "Run details are not available."),
      created,
      updated,
      deleted,
      errors: Number(row.errors ?? 0),
    };
  });
}

let activeCsrfToken = "";

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
  if (!response.ok) throw new ApiError((data as { error?: string }).error || `Request failed (${response.status})`, response.status);
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
  const [people, setPeople] = useState<Person[]>([]);
  const [runs, setRuns] = useState<Run[]>([]);
  const [selectedRun, setSelectedRun] = useState<Run | null>(null);
  const [syncRunning, setSyncRunning] = useState(false);
  const [lastSync, setLastSync] = useState("Never");
  const [health, setHealth] = useState("Setup required");
  const [counts, setCounts] = useState<{ users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number } | null>(null);
  const [resourceErrors, setResourceErrors] = useState({ people: false, runs: false });
  const [config, setConfig] = useState({
    schoolboxUrl: "",
    schoolboxJwt: "",
    serviceAccountJson: "",
    adminEmail: "",
    interval: "360",
    pastDays: "30",
    futureDays: "180",
    syncNewUsersByDefault: false,
    hasSchoolboxToken: false,
    hasGoogleServiceAccount: false,
    serviceAccountClientId: "",
  });

  const canOperate = Boolean(auth?.permissions.includes("operate"));
  const canConfigure = Boolean(auth?.permissions.includes("configure"));
  const canManageAccess = Boolean(auth?.permissions.includes("manage_access"));

  const acceptSession = useCallback((session: AuthSession) => {
    activeCsrfToken = session.csrfToken;
    setAuth(session);
    setView("dashboard");
  }, []);

  useEffect(() => {
    let cancelled = false;
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

  const loadData = useCallback(async () => {
    setLoading(true);
    setApiOnline(false);
    const results = await Promise.allSettled([
      fetchJson("/api/status"),
      fetchJson("/api/config"),
      fetchJson("/api/diagnostics"),
      fetchJson("/api/users"),
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
      const liveCounts = payload.counts as Record<string, unknown> | undefined;
      if (liveCounts) setCounts({ users: Number(liveCounts.users ?? 0), enabled: Number(liveCounts.enabled ?? liveCounts.users ?? 0), disabled: Number(liveCounts.disabled ?? 0), healthy: Number(liveCounts.healthy ?? 0), errors: Number(liveCounts.errors ?? 0), unmatched: Number(liveCounts.unmatched ?? 0), events: Number(liveCounts.events ?? 0) });
      setHealth(String(payload.health ?? payload.status ?? (!payload.configured ? "Setup required" : lastRun?.status === "failed" ? "Failed" : lastRun?.status === "completed_with_errors" ? "Warning" : "Healthy")));
      const syncDate = payload.lastSync ?? payload.last_sync_at ?? lastRun?.completedAt ?? lastRun?.startedAt;
      setLastSync(syncDate ? new Date(String(syncDate)).toLocaleString("en-AU") : "Never");
      if (statusConfig) setConfig(current => ({
        ...current,
        interval: String(statusConfig.syncIntervalMinutes ?? current.interval),
        pastDays: String(statusConfig.pastDays ?? current.pastDays),
        futureDays: String(statusConfig.futureDays ?? current.futureDays),
        syncNewUsersByDefault: Boolean(statusConfig.syncNewUsersByDefault ?? current.syncNewUsersByDefault),
      }));
      const fetchedPeople = normalisePeople(payload.people ?? payload.users);
      const fetchedRuns = normaliseRuns(payload.runs ?? payload.history ?? (payload.lastRun ? [payload.lastRun] : null));
      if (fetchedPeople) setPeople(fetchedPeople);
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
        syncNewUsersByDefault: Boolean(incoming.syncNewUsersByDefault ?? current.syncNewUsersByDefault),
        hasSchoolboxToken: Boolean(incoming.hasSchoolboxToken ?? current.hasSchoolboxToken),
        hasGoogleServiceAccount: Boolean(incoming.hasGoogleServiceAccount ?? current.hasGoogleServiceAccount),
        serviceAccountClientId: String(incoming.serviceAccountClientId ?? current.serviceAccountClientId),
      }));
    }
    if (diagnosticResult.status === "fulfilled") {
      const payload = diagnosticResult.value;
      const fetchedRuns = normaliseRuns(payload.runs ?? payload.history);
      if (fetchedRuns?.length) setRuns(fetchedRuns);
    }
    const usersResult = results[3];
    setResourceErrors({ people: usersResult.status === "rejected", runs: results[4].status === "rejected" });
    if (usersResult.status === "fulfilled") {
      const fetchedPeople = normalisePeople(usersResult.value.users);
      if (fetchedPeople) setPeople(fetchedPeople);
    }
    const runsResult = results[4];
    if (runsResult.status === "fulfilled") {
      const fetchedRuns = normaliseRuns(runsResult.value.runs);
      if (fetchedRuns) setRuns(fetchedRuns);
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    if (!auth) return;
    const timer = window.setTimeout(() => void loadData(), 0);
    return () => window.clearTimeout(timer);
  }, [auth, loadData]);
  useEffect(() => {
    if (!notice) return;
    const timer = window.setTimeout(() => setNotice(null), 5000);
    return () => window.clearTimeout(timer);
  }, [notice]);

  const changeView = (next: View) => {
    setView(next);
    setMobileNav(false);
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  const runNow = async () => {
    setSyncRunning(true);
    setNotice({ kind: "info", message: "Discovering users and syncing enabled calendars…" });
    try {
      const data = await fetchJson("/api/sync/run", { method: "POST", body: JSON.stringify({ trigger: "manual" }) });
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

  const saveConfig = async (message = "Settings saved") => {
    try {
      const saved = await fetchJson("/api/config", {
        method: "PUT",
        body: JSON.stringify({
          schoolboxBaseUrl: config.schoolboxUrl,
          ...(config.schoolboxJwt ? { schoolboxToken: config.schoolboxJwt } : {}),
          ...(config.serviceAccountJson ? { googleServiceAccountJson: config.serviceAccountJson } : {}),
          googleAdminEmail: config.adminEmail,
          syncIntervalMinutes: Number(config.interval),
          pastDays: Number(config.pastDays),
          futureDays: Number(config.futureDays),
          syncNewUsersByDefault: config.syncNewUsersByDefault,
          enabled: true,
          setupCompleted: true,
          timezone: "Australia/Sydney",
        }),
      });
      setApiOnline(true);
      setConfigured(Boolean(saved.schoolboxBaseUrl && saved.googleAdminEmail && saved.hasSchoolboxToken && saved.hasGoogleServiceAccount));
      setConfig(current => ({
        ...current,
        schoolboxJwt: "",
        serviceAccountJson: "",
        hasSchoolboxToken: Boolean(saved.hasSchoolboxToken ?? (current.hasSchoolboxToken || Boolean(current.schoolboxJwt))),
        hasGoogleServiceAccount: Boolean(saved.hasGoogleServiceAccount ?? (current.hasGoogleServiceAccount || Boolean(current.serviceAccountJson))),
        serviceAccountClientId: String(saved.serviceAccountClientId ?? (() => {
          try { return JSON.parse(current.serviceAccountJson || "{}").client_id ?? current.serviceAccountClientId; }
          catch { return current.serviceAccountClientId; }
        })()),
      }));
      setNotice({ kind: "success", message });
      return true;
    } catch (error) {
      if (error instanceof ApiError && error.status === 401) setAuth(null);
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "Settings could not be saved." });
      return false;
    }
  };

  const signOut = async () => {
    try {
      await fetchJson("/api/auth/logout", { method: "POST", body: "{}" });
      activeCsrfToken = "";
      setAuth(null);
      setView("dashboard");
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? `Sign out failed: ${error.message}` : "Sign out failed. Try again." });
    }
  };
  const handleSignedOut = useCallback(() => {
    activeCsrfToken = "";
    setAuth(null);
    setView("dashboard");
  }, []);

  if (auth === undefined) return <div className="auth-shell"><div className="auth-loading"><span className="brand-mark" aria-hidden="true"><i /><i /><i /></span><p>Starting Relay…</p></div></div>;
  if (!auth) return <LoginScreen readiness={readiness} unavailable={authUnavailable} onAuthenticated={acceptSession} />;

  const title = { dashboard: "Calendar operations", setup: "Connection setup", people: "People & sync coverage", runs: "Runs & troubleshooting", settings: "Sync settings", access: "IT access" }[view];
  const subtitle = {
    dashboard: configured ? "Monitor Schoolbox calendar delivery across Google Workspace." : "Complete setup before discovering users and starting calendar sync.",
    setup: "Connect both systems, grant access, then choose what Relay should sync.",
    people: "Review identity matches and choose whose calendars Relay maintains.",
    runs: "Inspect every sync and find the cause of exceptions.",
    settings: "Control schedule, calendar coverage and operational alerts.",
    access: "Configure Google sign-in and control who can administer Relay.",
  }[view];
  const initials = auth.displayName.split(/\s+/).map(part => part[0]).join("").slice(0, 2).toUpperCase() || "A";
  const roleLabel = auth.isOwner ? "Local administrator" : auth.role === "admin" ? "Administrator" : auth.role === "operator" ? "Operator" : "Viewer";

  return (
    <div className="app-shell">
      <aside className={`sidebar ${mobileNav ? "sidebar-open" : ""}`}>
        <button className="brand" onClick={() => changeView("dashboard")} aria-label="Relay home">
          <span className="brand-mark" aria-hidden="true"><i /><i /><i /></span>
          <span><b>Relay</b><small>Calendar operations</small></span>
        </button>
        <nav aria-label="Main navigation">
          <p className="nav-label">Workspace</p>
          <NavButton active={view === "dashboard"} icon="⌂" label="Overview" onClick={() => changeView("dashboard")} />
          {canConfigure && <NavButton active={view === "setup"} icon="↗" label="Setup" onClick={() => changeView("setup")} />}
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
            <div><p className="eyebrow">Schoolbox <span>→</span> Google Calendar</p><h1>{title}</h1><p>{subtitle}</p></div>
            {view !== "setup" && view !== "access" && <div className="heading-actions"><button className="button ghost" onClick={() => void loadData()} disabled={loading}>{loading ? "Refreshing…" : "Refresh"}</button>{canOperate && <button className="button primary" onClick={() => void runNow()} disabled={syncRunning || !configured}>{syncRunning ? "Starting…" : "Run sync now"}<span aria-hidden="true">→</span></button>}</div>}
          </div>

          {notice && <div role="status" className={`notice ${notice.kind}`}><span>{notice.kind === "success" ? "✓" : notice.kind === "error" ? "!" : "i"}</span>{notice.message}<button onClick={() => setNotice(null)} aria-label="Dismiss notification">×</button></div>}
          {view === "dashboard" && <Dashboard people={people} runs={runs} counts={counts} lastSync={lastSync} health={health} apiOnline={apiOnline} configured={configured} config={config} runsError={resourceErrors.runs} onNavigate={changeView} onSelectRun={(run) => { setSelectedRun(run); setView("runs"); }} />}
          {view === "setup" && canConfigure && <SetupWizard config={config} setConfig={setConfig} saveConfig={saveConfig} setNotice={setNotice} changeView={changeView} />}
          {view === "people" && <PeoplePage people={people} setPeople={setPeople} counts={counts} loadError={resourceErrors.people} canConfigure={canConfigure} setNotice={setNotice} />}
          {view === "runs" && <RunsPage runs={runs} selectedRun={selectedRun} setSelectedRun={setSelectedRun} runNow={runNow} syncRunning={syncRunning} canOperate={canOperate} loadError={resourceErrors.runs} />}
          {view === "settings" && canConfigure && <SettingsPage config={config} setConfig={setConfig} saveConfig={saveConfig} />}
          {view === "access" && canManageAccess && <AccessPage setNotice={setNotice} onSignedOut={handleSignedOut} />}
        </div>
      </main>
    </div>
  );
}

function LoginScreen({ readiness, unavailable, onAuthenticated }: { readiness: AuthReadiness; unavailable: boolean; onAuthenticated: (session: AuthSession) => void }) {
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
      {unavailable && <div className="auth-error" role="alert"><span>!</span><div>Relay could not reach its authentication service. <button onClick={() => window.location.reload()}>Try again</button></div></div>}
      {(error || callbackError) && <div className="auth-error" role="alert"><span>!</span>{error || callbackError}</div>}
      {!unavailable && readiness.localAdministrator ? <form onSubmit={submit} className="auth-form">
        <Field label="Username"><input autoComplete="username" value={username} onChange={event => setUsername(event.target.value)} required /></Field>
        <Field label="Password"><input type="password" autoComplete="current-password" value={password} onChange={event => setPassword(event.target.value)} required /></Field>
        <button className="button primary full" type="submit" disabled={busy}>{busy ? "Signing in…" : "Sign in as local administrator"}<span>→</span></button>
      </form> : !unavailable && <div className="auth-setup"><b>Local administrator setup is required</b><p>On the server, run <code>npm run auth:bootstrap</code> before opening Relay to the IT network.</p></div>}
      {!unavailable && readiness.googleSignInConfigured && <div className="oauth-choice"><span>or</span><a className="button google-signin" href="/api/auth/google/start"><b>G</b> Continue with Google Workspace</a><small>Only accounts pre-approved by the local administrator can enter.</small></div>}
      <div className="auth-foot"><span>🔒</span><p>Self-hosted on your internal server. Sessions expire after 30 minutes of inactivity.</p></div>
    </div>
  </div>;
}

function NavButton({ active, icon, label, count, onClick }: { active: boolean; icon: string; label: string; count?: string; onClick: () => void }) {
  return <button className={`nav-button ${active ? "active" : ""}`} onClick={onClick}><span className="nav-icon" aria-hidden="true">{icon}</span><span>{label}</span>{count && <small>{count}</small>}</button>;
}

function Dashboard({ people, runs, counts, lastSync, health, apiOnline, configured, config, runsError, onNavigate, onSelectRun }: { people: Person[]; runs: Run[]; counts: { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number } | null; lastSync: string; health: string; apiOnline: boolean; configured: boolean; config: Config; runsError: boolean; onNavigate: (view: View) => void; onSelectRun: (run: Run) => void }) {
  const enabledPeople = people.filter(person => person.syncEnabled);
  const totalUsers = people.length ? enabledPeople.length : counts?.enabled ?? 0;
  const healthyUsers = people.length ? enabledPeople.filter(person => person.status === "Synced" || person.status === "Syncing").length : counts?.healthy ?? 0;
  const issueUsers = people.length ? enabledPeople.filter(person => person.status === "Unmatched" || person.status === "Error").length : (counts?.errors ?? 0) + (counts?.unmatched ?? 0);
  const latestRun = runs[0];
  const activity = [...runs.slice(0, 7)].reverse();
  const activityMaximum = Math.max(1, ...activity.flatMap(run => [run.created ?? 0, run.updated ?? 0]));
  const attention = health.toLowerCase().includes("fail") || health.toLowerCase().includes("warning");
  const healthTitle = !apiOnline ? "Service unavailable" : !configured ? "Setup is not complete" : attention ? "Attention needed" : latestRun ? "Synchronization is healthy" : "Ready for the first sync";
  const healthTone = !apiOnline || attention ? "danger" : configured ? "success" : "warning";
  return <>
    <section className={`health-banner ${healthTone}`}>
      <div className="health-orbit"><span>{healthTone === "success" ? "✓" : "!"}</span></div>
      <div><p className="eyebrow">Current sync health</p><h2>{healthTitle}</h2><p>{lastSync === "Never" ? "No completed sync has been recorded." : `Last completed ${lastSync}.`}</p></div>
      <div className="health-meta"><span className={`status-pill ${healthTone}`}><i /> {apiOnline ? health : "Offline"}</span><small>Every {config.interval} minutes</small></div>
    </section>

    <section className="metric-grid" aria-label="Sync summary">
      <Metric label="People in sync" value={healthyUsers.toLocaleString()} detail={totalUsers ? `of ${totalUsers.toLocaleString()} enabled` : people.length ? "No users enabled" : "No users discovered"} delta={totalUsers ? `${(healthyUsers / totalUsers * 100).toFixed(1)}%` : "—"} />
      <Metric label="Calendar items" value={(counts?.events ?? 0).toLocaleString()} detail="inside active window" delta="Managed by Relay" />
      <Metric label="Last run" value={latestRun?.duration ?? "—"} detail={latestRun ? `${latestRun.users.toLocaleString()} enabled people synced` : "No run recorded"} delta={latestRun?.status ?? "Waiting"} />
      <Metric label="Needs attention" value={String(issueUsers)} detail="enabled people with errors" delta="Review" warning />
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
        <PanelHead title="Calendar coverage" subtitle="Schoolbox content included in every sync" />
        <div className="coverage-checks"><span><i>✓</i>Timetable lessons</span><span><i>✓</i>Resource bookings</span><span><i>✓</i>School and individual events</span></div>
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
  return <tr><td><button className="table-link" onClick={onClick}>{run.id}</button><small>{run.trigger}</small></td><td>{run.started}</td><td><StatusPill status={run.status} /></td><td>{run.users.toLocaleString()}<small className="cell-detail">of {run.usersDiscovered.toLocaleString()} discovered</small></td><td>{run.changes.toLocaleString()}</td><td>{run.duration}</td><td><button className="row-open" onClick={onClick} aria-label={`Open ${run.id}`}>→</button></td></tr>;
}

function StatusPill({ status }: { status: Person["status"] | Run["status"] }) {
  const tone = status === "Succeeded" || status === "Synced" || status === "Syncing" ? "success" : status === "Failed" || status === "Unmatched" || status === "Error" ? "danger" : status === "Running" ? "info" : "warning";
  return <span className={`status-pill ${tone}`}><i />{status}</span>;
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
  hasSchoolboxToken: boolean;
  hasGoogleServiceAccount: boolean;
  serviceAccountClientId: string;
};

function SetupWizard({ config, setConfig, saveConfig, setNotice, changeView }: { config: Config; setConfig: React.Dispatch<React.SetStateAction<Config>>; saveConfig: (message?: string) => Promise<boolean>; setNotice: (notice: Notice) => void; changeView: (view: View) => void }) {
  const [step, setStep] = useState(1);
  const [testing, setTesting] = useState<"schoolbox" | "google" | null>(null);
  const [schoolboxTested, setSchoolboxTested] = useState(false);
  const [googleTested, setGoogleTested] = useState(false);
  const [finished, setFinished] = useState(false);
  const clientId = useMemo(() => {
    try { const parsed = JSON.parse(config.serviceAccountJson || "{}"); return String(parsed.client_id ?? config.serviceAccountClientId); } catch { return config.serviceAccountClientId; }
  }, [config.serviceAccountJson, config.serviceAccountClientId]);

  const testConnection = async (target: "schoolbox" | "google") => {
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
    setTesting(target);
    try {
      const data = await fetchJson("/api/diagnostics", { method: "POST", body: JSON.stringify({ target, config }) });
      if (target === "schoolbox") setSchoolboxTested(true); else setGoogleTested(true);
      setNotice({ kind: "success", message: String(data.message ?? `${target === "schoolbox" ? "Schoolbox" : "Google Workspace"} connection verified.`) });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Connection test failed.";
      setNotice({ kind: "error", message });
    } finally { setTesting(null); }
  };

  const copyScopes = async () => { await navigator.clipboard.writeText(SCOPES.join(",")); setNotice({ kind: "success", message: "Required scopes copied to clipboard." }); };
  const finish = async () => {
    if (!schoolboxTested || !googleTested) {
      setNotice({ kind: "error", message: "Verify both Schoolbox and Google Workspace before activating Relay." });
      return;
    }
    if (await saveConfig("Setup complete. Relay is ready for its first sync.")) setFinished(true);
  };

  if (finished) return <section className="setup-complete"><div className="completion-mark">✓</div><p className="eyebrow">Connections complete</p><h2>Relay is ready to move.</h2><p>Your settings are saved. The first run will discover everyone and sync only users enabled by your coverage policy.</p><div className="finish-summary"><span><b>Schoolbox</b><small>{config.schoolboxUrl}</small></span><span><b>Google Workspace</b><small>{config.adminEmail}</small></span><span><b>New users</b><small>{config.syncNewUsersByDefault ? "Enabled automatically" : "Paused for review"}</small></span></div><button className="button primary" onClick={() => changeView("dashboard")}>Go to overview <span>→</span></button></section>;

  return <div className="setup-layout">
    <aside className="setup-steps" aria-label="Setup progress">
      <p className="nav-label">Configuration</p>
      {[
        [1, "Schoolbox", "Source connection"], [2, "Google service account", "Credentials"], [3, "Domain delegation", "Authorise access"], [4, "Verify connection", "Test permissions"], [5, "Sync policy", "Schedule & range"], [6, "Review", "Save configuration"],
      ].map(([number, name, desc]) => <button key={number} disabled={Number(number) > step} className={`${step === number ? "active" : ""} ${step > Number(number) ? "done" : ""}`} onClick={() => setStep(Number(number))}><span>{step > Number(number) ? "✓" : number}</span><div><b>{name}</b><small>{desc}</small></div></button>)}
      <div className="setup-help"><span>?</span><div><b>Need a hand?</b><small>Keep this screen open while you work through Google Admin.</small><a href="https://support.google.com/a/answer/162106" target="_blank" rel="noreferrer">Delegation guide ↗</a></div></div>
    </aside>

    <section className="setup-card">
      <div className="setup-progress"><span>Step {step} of 6</span><div><i style={{ width: `${step / 6 * 100}%` }} /></div></div>
      {step === 1 && <WizardSection eyebrow="Source connection" title="Connect your Schoolbox" intro="Relay reads each person’s timetable, events and due dates through the Schoolbox API. Your JWT is stored securely and never shown again.">
        <Field label="Schoolbox base URL" hint="The address your school uses to access Schoolbox."><div className="input-prefix"><span>https://</span><input value={config.schoolboxUrl.replace(/^https?:\/\//, "")} onChange={e => { setSchoolboxTested(false); setConfig(c => ({ ...c, schoolboxUrl: `https://${e.target.value}` })); }} placeholder="school.schoolbox.com.au" /></div></Field>
        <Field label="API JWT" hint={config.hasSchoolboxToken ? "A token is stored. Enter a value only to replace it." : "In Schoolbox Admin, edit the superuser, scroll to TOKENS, then choose Create token."}><input type="password" autoComplete="off" value={config.schoolboxJwt} onChange={e => { setSchoolboxTested(false); setConfig(c => ({ ...c, schoolboxJwt: e.target.value })); }} placeholder={config.hasSchoolboxToken ? "Stored securely" : "Paste your Schoolbox JWT"} /></Field>
        <div className="callout"><span>i</span><div><b>Give Relay read-only access</b><p>The token needs access to users, calendars, events and timetable data. It should not have permission to edit Schoolbox content.</p></div></div>
        <WizardActions><button className="button secondary" onClick={() => void testConnection("schoolbox")} disabled={testing !== null}>{testing === "schoolbox" ? "Testing…" : schoolboxTested ? "✓ Connection verified" : "Test connection"}</button><button className="button primary" onClick={() => setStep(2)} disabled={!schoolboxTested}>Continue <span>→</span></button></WizardActions>
      </WizardSection>}

      {step === 2 && <WizardSection eyebrow="Google credentials" title="Add a service account" intro="A Google Cloud service account lets Relay work quietly in the background—no individual sign-ins or calendar installs needed.">
        <ol className="instruction-list"><li><span>1</span><div><b>Open Google Cloud Console</b><p>Create or select a project for Relay, then enable the Google Calendar API and Admin SDK API.</p><a href="https://console.cloud.google.com/apis/library" target="_blank" rel="noreferrer">Open API Library ↗</a></div></li><li><span>2</span><div><b>Create a service account</b><p>Under IAM & Admin, create a service account and turn on domain-wide delegation.</p><a href="https://console.cloud.google.com/iam-admin/serviceaccounts" target="_blank" rel="noreferrer">Open service accounts ↗</a></div></li><li><span>3</span><div><b>Download a JSON key</b><p>Create a new JSON key for the account, then paste its complete contents below.</p></div></li></ol>
        <Field label="Service account JSON" hint={config.hasGoogleServiceAccount ? "A service account is stored. Enter JSON only to replace it." : "Relay encrypts this credential at rest."}><textarea rows={7} value={config.serviceAccountJson} onChange={e => { setGoogleTested(false); setConfig(c => ({ ...c, serviceAccountJson: e.target.value })); }} placeholder={config.hasGoogleServiceAccount ? "Stored securely" : '{\n  "type": "service_account",\n  "client_id": "..."\n}'} /></Field>
        {clientId && <div className="detected-value"><span>✓</span><div><b>Service account recognised</b><small>OAuth client ID: {clientId}</small></div></div>}
        <WizardActions><button className="button ghost" onClick={() => setStep(1)}>Back</button><button className="button primary" onClick={() => setStep(3)} disabled={!clientId}>Continue <span>→</span></button></WizardActions>
      </WizardSection>}

      {step === 3 && <WizardSection eyebrow="Google Admin" title="Grant domain-wide delegation" intro="Authorise the service account once for your whole organisation. No user passwords are shared with Relay.">
        <div className="delegation-box"><div className="delegation-number">1</div><div><h3>Open Google Admin Console</h3><p>Go to Security → Access and data control → API controls → Manage domain-wide delegation.</p><a className="button secondary" href="https://admin.google.com/ac/owl/domainwidedelegation" target="_blank" rel="noreferrer">Open Google Admin <span>↗</span></a></div></div>
        <div className="delegation-box"><div className="delegation-number">2</div><div><h3>Add a new API client</h3><p>Choose “Add new” and paste this numeric client ID:</p><CopyBox value={clientId || "Upload service account JSON in the previous step"} onCopy={() => navigator.clipboard.writeText(clientId)} /></div></div>
        <div className="delegation-box"><div className="delegation-number">3</div><div><h3>Authorise the required scopes</h3><p>Paste both scopes as a single comma-separated value, then choose Authorise.</p><div className="scope-list">{SCOPES.map(scope => <code key={scope}>{scope}</code>)}</div><button className="text-button" onClick={() => void copyScopes()}>Copy both scopes <span>□</span></button></div></div>
        <div className="callout warm"><span>!</span><div><b>Google may take a few minutes</b><p>Delegation changes can take up to 10 minutes to become active. If validation fails, wait briefly and try again.</p></div></div>
        <WizardActions><button className="button ghost" onClick={() => setStep(2)}>Back</button><button className="button primary" onClick={() => setStep(4)}>I’ve authorised it <span>→</span></button></WizardActions>
      </WizardSection>}

      {step === 4 && <WizardSection eyebrow="Permission check" title="Verify Google access" intro="Relay will impersonate a Workspace administrator to list users and manage their calendars. Use a dedicated admin account where possible.">
        <Field label="Delegated admin email" hint="This account must be active and have access to users across your Workspace domain."><input type="email" value={config.adminEmail} onChange={e => { setGoogleTested(false); setConfig(c => ({ ...c, adminEmail: e.target.value })); }} placeholder="calendar-admin@school.edu.au" /></Field>
        <div className={`validation-card ${googleTested ? "passed" : ""}`}><div className="validation-icon">{googleTested ? "✓" : "↻"}</div><div><h3>{googleTested ? "Google Workspace is ready" : "Run the access check"}</h3><p>{googleTested ? "Directory users are visible and Relay can write to a test calendar." : "We’ll test directory access, delegation and Calendar API permissions without changing user data."}</p></div><button className="button secondary" onClick={() => void testConnection("google")} disabled={testing !== null}>{testing === "google" ? "Checking…" : googleTested ? "Test again" : "Verify access"}</button></div>
        <div className="check-grid"><span className={googleTested ? "checked" : ""}>✓ <small>Service account</small></span><span className={googleTested ? "checked" : ""}>✓ <small>Directory access</small></span><span className={googleTested ? "checked" : ""}>✓ <small>Calendar access</small></span></div>
        <WizardActions><button className="button ghost" onClick={() => setStep(3)}>Back</button><button className="button primary" onClick={() => setStep(5)} disabled={!googleTested}>Continue <span>→</span></button></WizardActions>
      </WizardSection>}

      {step === 5 && <WizardSection eyebrow="Sync policy" title="Choose your coverage" intro="Set how often calendars update and how much history Relay maintains. These choices can be changed later.">
        <div className="form-grid"><Field label="Run every"><select value={config.interval} onChange={e => setConfig(c => ({ ...c, interval: e.target.value }))}><option value="15">15 minutes</option><option value="30">30 minutes</option><option value="60">60 minutes</option><option value="360">6 hours</option></select></Field><Field label="Past events"><select value={config.pastDays} onChange={e => setConfig(c => ({ ...c, pastDays: e.target.value }))}><option value="7">7 days</option><option value="14">14 days</option><option value="30">30 days</option><option value="90">90 days</option></select></Field><Field label="Future events"><select value={config.futureDays} onChange={e => setConfig(c => ({ ...c, futureDays: e.target.value }))}><option value="90">90 days</option><option value="180">180 days</option><option value="365">1 year</option></select></Field></div>
        <label className="policy-choice"><input type="checkbox" checked={config.syncNewUsersByDefault} onChange={e => setConfig(c => ({ ...c, syncNewUsersByDefault: e.target.checked }))} /><span><b>Automatically enable newly discovered users</b><small>{config.syncNewUsersByDefault ? "Future users will join calendar sync on their first discovery." : "Recommended for a pilot: new users are listed on People but remain paused until an administrator enables them."}</small></span></label>
        <div className="callout"><span>i</span><div><b>Complete Schoolbox calendar coverage</b><p>Relay includes timetable lessons, resource bookings, school events and individual events. Items managed by Relay are removed from Google when they disappear from Schoolbox.</p></div></div>
        <WizardActions><button className="button ghost" onClick={() => setStep(4)}>Back</button><button className="button primary" onClick={() => setStep(6)}>Review setup <span>→</span></button></WizardActions>
      </WizardSection>}

      {step === 6 && <WizardSection eyebrow="Ready to connect" title="Review your configuration" intro="Relay will save these connections and prepare user discovery and calendar sync.">
        <div className="review-list"><ReviewRow label="Schoolbox" value={config.schoolboxUrl} detail={schoolboxTested ? "Connection verified" : "Credentials supplied"} onEdit={() => setStep(1)} /><ReviewRow label="Google Workspace" value={config.adminEmail} detail={googleTested ? "Delegation verified" : "Delegated administrator"} onEdit={() => setStep(4)} /><ReviewRow label="Schedule" value={`Every ${config.interval} minutes`} detail={`${config.pastDays} days back · ${config.futureDays} days ahead`} onEdit={() => setStep(5)} /><ReviewRow label="New users" value={config.syncNewUsersByDefault ? "Enable automatically" : "Pause for review"} detail="Existing selections are managed from People" onEdit={() => setStep(5)} /></div>
        <div className="callout"><span>i</span><div><b>The first run discovers everyone</b><p>Only enabled users receive calendar changes. With the pilot setting, the first run safely populates People without writing events.</p></div></div>
        <WizardActions><button className="button ghost" onClick={() => setStep(5)}>Back</button><button className="button primary" onClick={() => void finish()}>Save and activate Relay <span>→</span></button></WizardActions>
      </WizardSection>}
    </section>
  </div>;
}

function WizardSection({ eyebrow, title, intro, children }: { eyebrow: string; title: string; intro: string; children: React.ReactNode }) { return <><p className="eyebrow">{eyebrow}</p><h2 className="setup-title">{title}</h2><p className="setup-intro">{intro}</p><div className="setup-content">{children}</div></>; }
function WizardActions({ children }: { children: React.ReactNode }) { return <div className="wizard-actions">{children}</div>; }
function Field({ label, hint, children }: { label: string; hint?: string; children: React.ReactNode }) { return <label className="field"><span>{label}</span>{children}{hint && <small>{hint}</small>}</label>; }
function CopyBox({ value, onCopy }: { value: string; onCopy: () => void }) { const [done, setDone] = useState(false); return <div className="copy-box"><code>{value}</code><button onClick={() => { onCopy(); setDone(true); window.setTimeout(() => setDone(false), 1500); }}>{done ? "Copied" : "Copy"}</button></div>; }
function ReviewRow({ label, value, detail, onEdit }: { label: string; value: string; detail: string; onEdit: () => void }) { return <div className="review-row"><span className="review-check">✓</span><div><small>{label}</small><b>{value || "Not supplied"}</b><p>{detail}</p></div><button onClick={onEdit}>Edit</button></div>; }

function PeoplePage({ people, setPeople, counts, loadError, canConfigure, setNotice }: {
  people: Person[];
  setPeople: React.Dispatch<React.SetStateAction<Person[]>>;
  counts: { users: number; enabled: number; disabled: number; healthy: number; errors: number; unmatched: number; events: number } | null;
  loadError: boolean;
  canConfigure: boolean;
  setNotice: (notice: Notice) => void;
}) {
  const [query, setQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState("All statuses");
  const [coverageFilter, setCoverageFilter] = useState("All coverage");
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [busy, setBusy] = useState(false);
  const [page, setPage] = useState(0);
  const pageSize = 100;
  const selectAllRef = useRef<HTMLInputElement>(null);
  const filtered = useMemo(() => people.filter(person => {
    const statusMatches = statusFilter === "All statuses" || person.status === statusFilter;
    const coverageMatches = coverageFilter === "All coverage" || (coverageFilter === "Enabled" ? person.syncEnabled : !person.syncEnabled);
    return statusMatches && coverageMatches && `${person.name} ${person.schoolboxEmail} ${person.googleEmail}`.toLowerCase().includes(query.toLowerCase());
  }), [people, query, statusFilter, coverageFilter]);
  const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
  const pageIndex = Math.min(page, pageCount - 1);
  const visible = filtered.slice(pageIndex * pageSize, (pageIndex + 1) * pageSize);
  const selectedIds = visible.filter(person => selected.has(person.id)).map(person => person.id);
  const selectedVisible = selectedIds.length;
  const enabledCount = people.length ? people.filter(person => person.syncEnabled).length : counts?.enabled ?? 0;
  const pausedCount = people.length ? people.length - enabledCount : counts?.disabled ?? 0;

  useEffect(() => {
    if (selectAllRef.current) selectAllRef.current.indeterminate = selectedVisible > 0 && selectedVisible < visible.length;
  }, [visible.length, selectedVisible]);

  const selectVisible = (checked: boolean) => {
    setSelected(checked ? new Set(visible.map(person => person.id)) : new Set());
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
        body: JSON.stringify({ userIds: ids, syncEnabled }),
      });
      const updated = Number(payload.updated ?? ids.length);
      const changed = new Set(ids);
      setPeople(current => current.map(person => changed.has(person.id) ? { ...person, syncEnabled } : person));
      setSelected(new Set());
      setNotice({
        kind: "success",
        message: `Calendar sync ${syncEnabled ? "enabled" : "paused"} for ${updated} ${updated === 1 ? "person" : "people"}. The change applies on the next run.`,
      });
    } catch (error) {
      setNotice({ kind: "error", message: error instanceof Error ? error.message : "User sync coverage could not be updated." });
    } finally {
      setBusy(false);
    }
  };
  const exportCsv = () => {
    const escape = (value: string) => `"${value.replaceAll('"', '""')}"`;
    const rows = [["Name", "Schoolbox email", "Google email", "Role", "Calendar sync", "Status", "Last sync"], ...filtered.map(person => [person.name, person.schoolboxEmail, person.googleEmail, person.role, person.syncEnabled ? "Enabled" : "Paused", person.status, person.lastSync])];
    const url = URL.createObjectURL(new Blob([rows.map(row => row.map(escape).join(",")).join("\n")], { type: "text/csv;charset=utf-8" }));
    const link = document.createElement("a");
    link.href = url;
    link.download = "relay-user-mappings.csv";
    link.click();
    URL.revokeObjectURL(url);
  };
  return <>
    <section className="people-summary"><div><span className="summary-icon green">#</span><p><b>{people.length || counts?.users || 0}</b><small>Discovered</small></p></div><div><span className="summary-icon green">✓</span><p><b>{enabledCount}</b><small>Enabled</small></p></div><div><span className="summary-icon amber">Ⅱ</span><p><b>{pausedCount}</b><small>Paused</small></p></div><div><span className="summary-icon red">!</span><p><b>{people.length ? people.filter(p => p.status === "Unmatched" || p.status === "Error").length : (counts?.unmatched ?? 0) + (counts?.errors ?? 0)}</b><small>Needs attention</small></p></div></section>
    <section className="panel people-panel" aria-busy={busy}><div className="people-tools"><div className="search-box"><span aria-hidden="true">⌕</span><input value={query} onChange={e => { setQuery(e.target.value); setPage(0); setSelected(new Set()); }} placeholder="Search people or email…" aria-label="Search people" /></div><select value={coverageFilter} onChange={e => { setCoverageFilter(e.target.value); setPage(0); setSelected(new Set()); }} aria-label="Filter calendar sync coverage"><option>All coverage</option><option>Enabled</option><option>Paused</option></select><select value={statusFilter} onChange={e => { setStatusFilter(e.target.value); setPage(0); setSelected(new Set()); }} aria-label="Filter sync status"><option>All statuses</option><option>Synced</option><option>Syncing</option><option>Pending</option><option>Unmatched</option><option>Error</option></select><button className="button ghost" onClick={exportCsv}>Export CSV</button></div>
      {canConfigure && selectedVisible > 0 && <div className="people-bulk" role="status"><b>{selectedVisible} selected</b><span>Bulk changes apply only to the selected visible users.</span><button className="button secondary" onClick={() => void updateCoverage(selectedIds, true)} disabled={busy}>Enable selected</button><button className="button ghost" onClick={() => void updateCoverage(selectedIds, false)} disabled={busy}>Pause selected</button></div>}
      <div className="coverage-note"><span>i</span><p><b>Pausing stops future updates.</b> Existing Relay-created Google events are left in place. Change the default for newly discovered users under <strong>Settings → People</strong>.</p></div>
      <div className="table-wrap"><table className="people-table"><caption className="sr-only">Discovered Google Workspace users and their Schoolbox calendar sync coverage</caption><thead><tr>{canConfigure && <th scope="col" className="selection-column"><input ref={selectAllRef} type="checkbox" checked={visible.length > 0 && selectedVisible === visible.length} onChange={event => selectVisible(event.target.checked)} aria-label="Select visible users" disabled={busy || visible.length === 0} /></th>}<th scope="col">Person</th><th scope="col">Schoolbox identity</th><th scope="col">Google Workspace</th><th scope="col">Role</th><th scope="col">Calendar sync</th><th scope="col">Status</th><th scope="col">Last sync</th></tr></thead><tbody>{visible.map(person => <tr key={person.id}>{canConfigure && <td className="selection-column"><input type="checkbox" checked={selected.has(person.id)} onChange={event => selectOne(person.id, event.target.checked)} aria-label={`Select ${person.name}`} disabled={busy} /></td>}<th scope="row" className="person-row-header"><div className="person-cell"><span className="person-avatar">{person.name.split(" ").map(part => part[0]).join("").slice(0, 2)}</span><div><b>{person.name}</b><small>{person.id}</small></div></div></th><td>{person.schoolboxEmail}</td><td className={person.googleEmail === "—" ? "muted" : ""}>{person.googleEmail}</td><td>{person.role}</td><td>{canConfigure ? <label className="sync-switch"><input type="checkbox" checked={person.syncEnabled} onChange={event => void updateCoverage([person.id], event.target.checked)} disabled={busy} aria-label={`Sync calendar for ${person.name}`} /><span aria-hidden="true" /><b>{person.syncEnabled ? "Enabled" : "Paused"}</b></label> : <span className={`coverage-state ${person.syncEnabled ? "enabled" : "paused"}`}>{person.syncEnabled ? "Enabled" : "Paused"}</span>}</td><td><StatusPill status={person.status} /></td><td>{person.lastSync}</td></tr>)}</tbody></table>{filtered.length === 0 && <div className="empty-state"><b>{loadError ? "People could not be loaded" : people.length === 0 ? "No people discovered yet" : "No people found"}</b><p>{loadError ? "Refresh after checking the server connection and logs." : people.length === 0 ? "Complete setup and run a sync to discover Workspace users. If the new-user default is paused, discovery will not write calendar events." : "Try different search or filter options."}</p></div>}</div>
      <div className="table-footer"><span>{filtered.length ? `Showing ${pageIndex * pageSize + 1}–${Math.min((pageIndex + 1) * pageSize, filtered.length)} of ${filtered.length} matching people` : `0 of ${people.length} people`}</span><div><button onClick={() => { setPage(Math.max(0, pageIndex - 1)); setSelected(new Set()); }} disabled={pageIndex === 0}>Previous</button><span>Page {pageIndex + 1} of {pageCount}</span><button onClick={() => { setPage(Math.min(pageCount - 1, pageIndex + 1)); setSelected(new Set()); }} disabled={pageIndex >= pageCount - 1}>Next</button></div></div>
    </section>
  </>;
}

function RunsPage({ runs, selectedRun, setSelectedRun, runNow, syncRunning, canOperate, loadError }: { runs: Run[]; selectedRun: Run | null; setSelectedRun: (run: Run | null) => void; runNow: () => Promise<void>; syncRunning: boolean; canOperate: boolean; loadError: boolean }) {
  const [status, setStatus] = useState("All statuses");
  const filtered = runs.filter(run => status === "All statuses" || run.status === status);
  const downloadRun = (run: Run) => {
    const url = URL.createObjectURL(new Blob([JSON.stringify(run, null, 2)], { type: "application/json" }));
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
    {selectedRun && <aside className="run-drawer"><div className="drawer-head"><div><p className="eyebrow">Run detail</p><h2>{selectedRun.id}</h2></div><button onClick={() => setSelectedRun(null)} aria-label="Close run details">×</button></div><StatusPill status={selectedRun.status} /><div className="drawer-summary"><span><small>Started</small><b>{selectedRun.started}</b></span><span><small>Duration</small><b>{selectedRun.duration}</b></span><span><small>Trigger</small><b>{selectedRun.trigger}</b></span><span><small>People synced</small><b>{selectedRun.users}</b></span></div><h3>Run narrative</h3><p>{selectedRun.note}</p><div className="event-stats"><div><span className="stat-mark green">+</span><b>{selectedRun.created ?? 0}</b><small>Created</small></div><div><span className="stat-mark navy">↻</span><b>{selectedRun.updated ?? 0}</b><small>Updated</small></div><div><span className="stat-mark red">−</span><b>{selectedRun.deleted ?? 0}</b><small>Removed</small></div></div><h3>Timeline</h3><ol className="run-timeline"><li className="done"><span>✓</span><div><b>Users discovered</b><small>{selectedRun.usersDiscovered} directory identities loaded</small></div></li><li className="done"><span>✓</span><div><b>Identity matches</b><small>{selectedRun.usersMatched} Schoolbox identities matched; {selectedRun.users} enabled calendars processed</small></div></li><li className={selectedRun.status === "Failed" ? "failed" : "done"}><span>{selectedRun.status === "Failed" ? "!" : "✓"}</span><div><b>Google calendars updated</b><small>{selectedRun.status === "Failed" ? "Stopped after an API error" : `${selectedRun.errors ?? 0} user errors recorded`}</small></div></li></ol>{canOperate && selectedRun.status !== "Succeeded" && <button className="button primary full" onClick={() => void runNow()}>Retry enabled-user sync <span>→</span></button>}<button className="button ghost full" onClick={() => downloadRun(selectedRun)}>Download run summary</button></aside>}
  </div>;
}

function SettingsPage({ config, setConfig, saveConfig }: { config: Config; setConfig: React.Dispatch<React.SetStateAction<Config>>; saveConfig: (message?: string) => Promise<boolean> }) {
  const [section, setSection] = useState("Schedule");
  const submit = (event: FormEvent) => { event.preventDefault(); void saveConfig(); };
  return <div className="settings-layout"><aside className="settings-nav">{["Schedule", "People", "Calendar content", "Connections", "Advanced"].map(item => <button key={item} className={section === item ? "active" : ""} onClick={() => setSection(item)}>{item}<span>→</span></button>)}</aside><form className="panel settings-card" onSubmit={submit}>
    {section === "Schedule" && <SettingsSection title="Sync schedule" intro="Choose how frequently Relay checks Schoolbox and the calendar window it maintains."><div className="form-grid"><Field label="Frequency"><select value={config.interval} onChange={e => setConfig(c => ({ ...c, interval: e.target.value }))}><option value="15">Every 15 minutes</option><option value="30">Every 30 minutes</option><option value="60">Every hour</option><option value="360">Every 6 hours</option></select></Field><Field label="Keep past events"><select value={config.pastDays} onChange={e => setConfig(c => ({ ...c, pastDays: e.target.value }))}><option value="7">7 days</option><option value="14">14 days</option><option value="30">30 days</option></select></Field><Field label="Sync ahead"><select value={config.futureDays} onChange={e => setConfig(c => ({ ...c, futureDays: e.target.value }))}><option value="90">90 days</option><option value="180">180 days</option><option value="365">1 year</option></select></Field></div><div className="settings-note"><span>◷</span><div><b>Next run follows the configured interval</b><small>Runs use the Australia/Sydney time zone.</small></div></div></SettingsSection>}
    {section === "People" && <SettingsSection title="New-user coverage" intro="Choose what happens the first time Relay discovers a Google Workspace user."><label className="policy-choice"><input type="checkbox" checked={config.syncNewUsersByDefault} onChange={e => setConfig(c => ({ ...c, syncNewUsersByDefault: e.target.checked }))} /><span><b>Automatically enable newly discovered users</b><small>{config.syncNewUsersByDefault ? "Newly matched people will be included in the next calendar sync." : "New users will appear under People as paused until an administrator enables them."}</small></span></label><div className="callout warm"><span>i</span><div><b>This changes future discoveries only</b><p>Existing user selections stay as they are. Use People to enable or pause individuals or a selected group.</p></div></div></SettingsSection>}
    {section === "Calendar content" && <SettingsSection title="Calendar content" intro="Relay mirrors the complete calendar feed Schoolbox exposes for each user."><div className="callout"><span>✓</span><div><b>Everything relevant is included</b><p>Timetable lessons, rooms, resource bookings, school events and individual events are synced. Relay only removes Google events it owns.</p></div></div></SettingsSection>}
    {section === "Connections" && <SettingsSection title="Connected services" intro="Review the identities Relay uses. Secrets are encrypted and cannot be revealed."><div className="connection-row"><span className="connection-logo">S</span><div><b>Schoolbox</b><small>{config.schoolboxUrl || "Complete Setup to connect"}</small></div><span className={`status-pill ${config.schoolboxUrl ? "success" : "warning"}`}><i />{config.schoolboxUrl ? "Connected" : "Not configured"}</span></div><div className="connection-row"><span className="connection-logo google">G</span><div><b>Google Workspace</b><small>{config.adminEmail || "Complete Setup to delegate access"}</small></div><span className={`status-pill ${config.adminEmail ? "success" : "warning"}`}><i />{config.adminEmail ? "Delegated" : "Not configured"}</span></div><div className="callout warm"><span>!</span><div><b>Rotate credentials from Setup</b><p>Re-run the connection tests before saving replacements. Existing Google Calendar events remain untouched.</p></div></div></SettingsSection>}
    {section === "Advanced" && <SettingsSection title="Advanced controls" intro="Operational safeguards for large organisations."><div className="callout"><span>✓</span><div><b>Safeguards are always active</b><p>Temporary API failures use exponential backoff, and Relay never modifies manually-created Google events.</p></div></div></SettingsSection>}
    <div className="settings-actions"><span>Changes take effect on the next run.</span><button className="button primary" type="submit">Save changes</button></div>
  </form></div>;
}

function AccessPage({ setNotice, onSignedOut }: { setNotice: (notice: Notice) => void; onSignedOut: () => void }) {
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
      <div className="settings-head"><p className="eyebrow">Owner only</p><h2>Google Workspace sign-in</h2><p>Create an Internal Web OAuth client in Google Cloud, then enter it here. This is separate from the service account used for calendar synchronization.</p></div>
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
      <div className="role-grid"><div><b>Viewer</b><small>Dashboard, people, and run history</small></div><div><b>Operator</b><small>Viewer access plus diagnostics and manual syncs</small></div><div><b>Administrator</b><small>Operator access plus connections and sync settings</small></div></div>
      <form className="staff-add-form" onSubmit={addStaff}><Field label="Google Workspace email"><input type="email" value={newStaff.email} onChange={event => setNewStaff(current => ({ ...current, email: event.target.value }))} placeholder="it.staff@school.edu.au" required /></Field><Field label="Display name"><input value={newStaff.displayName} onChange={event => setNewStaff(current => ({ ...current, displayName: event.target.value }))} placeholder="Optional" /></Field><Field label="Role"><select value={newStaff.role} onChange={event => setNewStaff(current => ({ ...current, role: event.target.value as StaffRole }))}><option value="viewer">Viewer</option><option value="operator">Operator</option><option value="admin">Administrator</option></select></Field><button className="button primary" type="submit" disabled={busy === "new-staff"}>{busy === "new-staff" ? "Adding…" : "Add staff"}</button></form>
      <div className="staff-list">{staff.map(account => <div className="staff-row" key={account.id}><span className="person-avatar">{(account.displayName || account.email).split(/\s|@/).map(part => part[0]).join("").slice(0, 2).toUpperCase()}</span><div className="staff-identity"><b>{account.displayName || account.email}</b><small>{account.email} · {account.linked ? `Linked · Last login ${account.lastLoginAt ? new Date(account.lastLoginAt).toLocaleString("en-AU") : "not recorded"}` : "Awaiting first Google sign-in"}</small></div><select value={account.role} onChange={event => updateStaff(account.id, { role: event.target.value as StaffRole })} aria-label={`Role for ${account.email}`}><option value="viewer">Viewer</option><option value="operator">Operator</option><option value="admin">Administrator</option></select><label className="enable-control"><input type="checkbox" checked={account.enabled} onChange={event => updateStaff(account.id, { enabled: event.target.checked })} />Enabled</label><button className="button secondary" onClick={() => void saveStaff(account)} disabled={busy === account.id}>Save</button><button className="row-delete" onClick={() => void removeStaff(account)} disabled={busy === account.id}>Remove</button></div>)}{staff.length === 0 && <div className="empty-state"><b>No Google Workspace staff added</b><p>Add an email address above. The account remains blocked until explicitly listed here.</p></div>}</div>
    </section>

    <section className="panel access-panel narrow-panel">
      <div className="settings-head"><p className="eyebrow">Break-glass account</p><h2>Local administrator password</h2><p>Changing this password signs out the current session. Keep the credential in your IT password vault.</p></div>
      <form className="password-form" onSubmit={changePassword}><Field label="Current password"><input type="password" autoComplete="current-password" value={passwords.currentPassword} onChange={event => setPasswords(current => ({ ...current, currentPassword: event.target.value }))} required /></Field><Field label="New password" hint="At least 14 characters."><input type="password" autoComplete="new-password" minLength={14} value={passwords.nextPassword} onChange={event => setPasswords(current => ({ ...current, nextPassword: event.target.value }))} required /></Field><Field label="Confirm new password"><input type="password" autoComplete="new-password" minLength={14} value={passwords.confirmation} onChange={event => setPasswords(current => ({ ...current, confirmation: event.target.value }))} required /></Field><button className="button secondary" type="submit" disabled={busy === "password"}>{busy === "password" ? "Changing…" : "Change password"}</button></form>
    </section>
  </div>;
}

function SettingsSection({ title, intro, children }: { title: string; intro: string; children: React.ReactNode }) { return <><div className="settings-head"><h2>{title}</h2><p>{intro}</p></div><div className="settings-content">{children}</div></>; }
