import {
  createContentHash,
  createDeterministicEventId,
  GoogleApiError,
  GoogleWorkspaceClient,
  parseServiceAccountJson,
  type GoogleCalendarEventInput,
  type GoogleDirectoryUser,
} from "./google";
import { SchoolboxClient, type NormalizedSchoolboxCalendarEvent, type SchoolboxUser } from "./schoolbox";
import {
  MicrosoftGraphClient,
  MicrosoftGraphError,
  type MicrosoftEventInput,
  type MicrosoftGraphUser,
} from "./microsoft";
import {
  eventExcludedForUser,
  eventIncludedByPolicy,
  resolveGoogleEventRule,
  withoutManagedCalendarDestination,
  type SyncPolicy,
} from "./policy";
import {
  addAudit,
  checkpointRun,
  checkpointRunTarget,
  createRun,
  deleteCalendarTargetRecords,
  deleteEventMapping,
  discoverUserMappings,
  finishRun,
  finishRunTarget,
  finishRunUserDiagnostic,
  getConfig,
  getEventMappings,
  getUserEventExclusions,
  getUserCalendarTarget,
  getUserMapping,
  listCalendarTargetsForDestination,
  listRunUserDiagnostics,
  listUserCalendarTargets,
  listRuns,
  recordManagedEventCleanup,
  recordRunEventDiagnostic,
  recoverStaleRuns,
  recordDiscoveredEventTypes,
  saveConfig,
  setUsersSyncEnabled,
  startRunTarget,
  startRunUserDiagnostic,
  touchRunHeartbeat,
  touchEventMapping,
  upsertEventMapping,
  upsertUserCalendarTarget,
  upsertUserMapping,
  type EventMapping,
  type RunSummary,
  type RunTargetSummary,
  type TargetProvider,
  type UserCalendarTarget,
} from "./storage";
import { HttpError } from "./security";

type MatchedUser = { google: GoogleDirectoryUser; schoolbox: SchoolboxUser; schoolboxEmail: string };
type MicrosoftMatchedUser = { microsoft: MicrosoftGraphUser; schoolbox: SchoolboxUser; schoolboxEmail: string };
type SchoolboxSyncClient = Pick<SchoolboxClient, "getAllUsers" | "getCalendarEvents">;
type GoogleSyncClient = Pick<GoogleWorkspaceClient, "listAllUsers" | "createCalendar" | "updateCalendar" | "insertEvent" | "updateEvent" | "deleteEvent">;
type GoogleCleanupClient = Pick<GoogleWorkspaceClient, "deleteEvent"> & Partial<Pick<GoogleWorkspaceClient, "deleteCalendar">>;
type GoogleCalendarRetirementClient = Pick<GoogleWorkspaceClient, "deleteCalendar">;
type MicrosoftCalendarRetirementClient = Pick<MicrosoftGraphClient, "deleteCalendar">;
type CalendarRetirementClient = GoogleCalendarRetirementClient | MicrosoftCalendarRetirementClient;
type MicrosoftSyncClient = Pick<MicrosoftGraphClient, "listAllUsers" | "createCalendar" | "updateCalendar" | "insertEvent" | "updateEvent" | "deleteEvent">;
type MicrosoftCleanupClient = Pick<MicrosoftGraphClient, "deleteEvent"> & Partial<Pick<MicrosoftGraphClient, "deleteCalendar">>;
type TargetCleanupClient = GoogleCleanupClient | MicrosoftCleanupClient;

/** Optional client overrides used by deterministic integration tests. */
export type SyncClientOverrides = {
  schoolbox?: SchoolboxSyncClient;
  google?: GoogleSyncClient;
  microsoft?: MicrosoftSyncClient;
};

/** Millisecond overrides used only by deterministic timeout tests. */
export type SyncRuntimeOptions = {
  discoveryTimeoutMs?: number;
  userSyncTimeoutMs?: number;
  runTimeoutMs?: number;
};

export type ManagedEventCleanupResult = {
  paused: true;
  deleted: number;
  alreadyMissing: number;
  remaining: number;
  calendarsDeleted: number;
  calendarsAlreadyMissing: number;
  calendarsRemaining: number;
  error: string | null;
};

export type CalendarDestinationRetirementResult = {
  destinationId: string;
  calendarsDeleted: number;
  calendarsAlreadyMissing: number;
  calendarsFailed: number;
  calendarsRemaining: number;
  eventMappingsRemoved: number;
  error: string | null;
};

class SyncTimeoutError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SyncTimeoutError";
  }
}

function durationLabel(milliseconds: number): string {
  if (milliseconds % 60_000 === 0) {
    const minutes = milliseconds / 60_000;
    return `${minutes} ${minutes === 1 ? "minute" : "minutes"}`;
  }
  const seconds = Math.max(1, Math.round(milliseconds / 1_000));
  return `${seconds} ${seconds === 1 ? "second" : "seconds"}`;
}

function positiveTimeout(override: number | undefined, configured: number): number {
  return override !== undefined && Number.isFinite(override) && override > 0
    ? override
    : configured;
}

function abortError(signal: AbortSignal, fallback: string): Error {
  return signal.reason instanceof Error ? signal.reason : new Error(fallback);
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw abortError(signal, "Calendar synchronization was aborted");
}

function createDeadlineSignal(
  parent: AbortSignal | undefined,
  milliseconds: number,
  message: string,
): { signal: AbortSignal; dispose: () => void } {
  const controller = new AbortController();
  const forwardParentAbort = () => controller.abort(parent?.reason ?? new Error("Calendar synchronization was aborted"));
  if (parent?.aborted) forwardParentAbort();
  else parent?.addEventListener("abort", forwardParentAbort, { once: true });
  const timer = setTimeout(() => controller.abort(new SyncTimeoutError(message)), milliseconds);
  return {
    signal: controller.signal,
    dispose: () => {
      clearTimeout(timer);
      parent?.removeEventListener("abort", forwardParentAbort);
    },
  };
}

async function withHardDeadline<T>(
  parent: AbortSignal | undefined,
  milliseconds: number,
  message: string,
  operation: (signal: AbortSignal) => Promise<T>,
): Promise<T> {
  const deadline = createDeadlineSignal(parent, milliseconds, message);
  let rejectAbort: (() => void) | undefined;
  const aborted = new Promise<never>((_resolve, reject) => {
    rejectAbort = () => reject(abortError(deadline.signal, message));
    if (deadline.signal.aborted) rejectAbort();
    else deadline.signal.addEventListener("abort", rejectAbort, { once: true });
  });
  try {
    return await Promise.race([operation(deadline.signal), aborted]);
  } finally {
    if (rejectAbort) deadline.signal.removeEventListener("abort", rejectAbort);
    deadline.dispose();
  }
}

function syncErrorMessage(error: unknown, signal?: AbortSignal): string {
  if (signal?.aborted) return abortError(signal, "Calendar synchronization was aborted").message;
  if (error instanceof MicrosoftGraphError) {
    const diagnostics = [
      `Graph code ${error.code}`,
      error.requestId ? `request ${error.requestId}` : null,
      error.date ? `at ${error.date}` : null,
      error.retryAfterMs !== undefined ? `retry after ${Math.ceil(error.retryAfterMs / 1_000)}s` : null,
    ].filter(Boolean).join(", ");
    return `${error.message}${diagnostics ? ` (${diagnostics})` : ""}`;
  }
  return error instanceof Error ? error.message : "Unknown synchronization error";
}

function normalizedEmail(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
}

/**
 * Index active Schoolbox users by either primary or alternate email.
 * A unique primary match takes precedence. Alternate emails are considered
 * only where no primary match exists, and ambiguous addresses at either level
 * are omitted so Relay cannot associate the wrong Schoolbox identity.
 */
export function indexActiveSchoolboxUsersByEmail(users: SchoolboxUser[]): Map<string, SchoolboxUser> {
  const primary = new Map<string, Map<number, SchoolboxUser>>();
  const alternate = new Map<string, Map<number, SchoolboxUser>>();
  for (const user of users.filter(isSchoolboxActive)) {
    for (const [source, value] of [[primary, user.email], [alternate, user.altEmail]] as const) {
      const email = normalizedEmail(value);
      if (!email) continue;
      const usersForEmail = source.get(email) ?? new Map<number, SchoolboxUser>();
      usersForEmail.set(user.id, user);
      source.set(email, usersForEmail);
    }
  }

  const unique = new Map<string, SchoolboxUser>();
  const emails = new Set([...primary.keys(), ...alternate.keys()]);
  for (const email of emails) {
    const primaryUsers = primary.get(email);
    if (primaryUsers?.size === 1) {
      unique.set(email, primaryUsers.values().next().value!);
      continue;
    }
    if (primaryUsers && primaryUsers.size > 1) continue;
    const alternateUsers = alternate.get(email);
    if (alternateUsers?.size === 1) unique.set(email, alternateUsers.values().next().value!);
  }
  return unique;
}

function schoolboxDisplayName(user: SchoolboxUser): string {
  const record = user as SchoolboxUser & {
    fullName?: string;
    firstName?: string;
    preferredName?: string;
    lastName?: string;
  };
  return record.fullName || [record.preferredName || record.firstName, record.lastName].filter(Boolean).join(" ") || user.email || `Schoolbox user ${user.id}`;
}

function googleDisplayName(user: GoogleDirectoryUser): string {
  const record = user as GoogleDirectoryUser & { name?: { fullName?: string } };
  return record.name?.fullName || user.primaryEmail;
}

function schoolboxRole(user: SchoolboxUser): string | null {
  const record = user as SchoolboxUser & { role?: { name?: string; type?: string } | string };
  return typeof record.role === "string" ? record.role : record.role?.name || record.role?.type || null;
}

function isSchoolboxActive(user: SchoolboxUser): boolean {
  const record = user as SchoolboxUser & { enabled?: boolean; isDeleted?: boolean };
  return record.enabled !== false && record.isDeleted !== true;
}

function isGoogleActive(user: GoogleDirectoryUser): boolean {
  const record = user as GoogleDirectoryUser & { suspended?: boolean; archived?: boolean };
  return !record.suspended && !record.archived;
}

function microsoftEmail(user: MicrosoftGraphUser): string {
  return normalizedEmail(user.mail || user.userPrincipalName);
}

function googleMatchEmails(user: GoogleDirectoryUser): string[] {
  return [...new Set([
    user.primaryEmail,
    ...(user.aliases ?? []),
    ...(user.nonEditableAliases ?? []),
  ].map(normalizedEmail).filter(Boolean))];
}

function microsoftMatchEmails(user: MicrosoftGraphUser): string[] {
  return [...new Set([
    user.mail ?? "",
    user.userPrincipalName ?? "",
    ...(user.proxyAddresses ?? []).map((address) => address.replace(/^smtp:/i, "")),
  ].map(normalizedEmail).filter(Boolean))];
}

function directorySchoolboxMatch(
  emails: string[],
  schoolboxByEmail: Map<string, SchoolboxUser>,
): { schoolbox: SchoolboxUser; matchedEmail: string } | null {
  for (const email of emails) {
    const schoolbox = schoolboxByEmail.get(email);
    if (schoolbox) return { schoolbox, matchedEmail: email };
  }
  return null;
}

function microsoftDisplayName(user: MicrosoftGraphUser): string {
  return user.displayName?.trim() || microsoftEmail(user);
}

function isMicrosoftActive(user: MicrosoftGraphUser): boolean {
  return user.accountEnabled !== false && user.userType?.toLowerCase() !== "guest" && Boolean(microsoftEmail(user));
}

function uuidFromHex(hex: string): string {
  const value = hex.slice(0, 32).padEnd(32, "0").split("");
  value[12] = "4";
  value[16] = ["8", "9", "a", "b"][Number.parseInt(value[16] ?? "0", 16) % 4];
  const joined = value.join("");
  return `${joined.slice(0, 8)}-${joined.slice(8, 12)}-${joined.slice(12, 16)}-${joined.slice(16, 20)}-${joined.slice(20, 32)}`;
}

function microsoftWallClock(value: string, timezone: string): string {
  const trimmed = value.trim();
  if (!/(?:Z|[+-]\d{2}:\d{2})$/i.test(trimmed)) return trimmed.replace(/\.\d{1,7}$/, "");
  const date = new Date(trimmed);
  if (!Number.isFinite(date.getTime())) return trimmed;
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: timezone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date);
  const valueFor = (type: Intl.DateTimeFormatPartTypes) => parts.find((part) => part.type === type)?.value ?? "";
  return `${valueFor("year")}-${valueFor("month")}-${valueFor("day")}T${valueFor("hour")}:${valueFor("minute")}:${valueFor("second")}`;
}

export async function microsoftEventBody(
  event: NormalizedSchoolboxCalendarEvent,
  microsoftUserId: string,
  timezone: string,
  sourceKey: string,
  policy: SyncPolicy,
  destinationId = "primary",
): Promise<MicrosoftEventInput> {
  const rule = resolveGoogleEventRule({ category: event.category ?? "other", type: event.type }, policy);
  const descriptionParts = [
    policy.includeDescription ? event.description : undefined,
    policy.includeEventTypeInDescription && event.type ? `Schoolbox type: ${event.type}` : undefined,
    policy.includeAuthorInDescription && event.author ? `Schoolbox author: ${event.author}` : undefined,
    policy.includeSchoolboxLink && event.sourceUrl ? `Schoolbox: ${event.sourceUrl}` : undefined,
  ].filter(Boolean);
  const transactionHash = await createContentHash({ provider: "microsoft", microsoftUserId, sourceKey, destinationId });
  const body: MicrosoftEventInput = {
    subject: `${policy.titlePrefix ? `${policy.titlePrefix} ` : ""}${event.title || "Schoolbox event"}`,
    body: { contentType: "text", content: descriptionParts.join("\n\n") },
    start: { dateTime: event.allDay ? `${event.start.slice(0, 10)}T00:00:00` : microsoftWallClock(event.start, timezone), timeZone: timezone },
    end: { dateTime: event.allDay ? `${event.end.slice(0, 10)}T00:00:00` : microsoftWallClock(event.end, timezone), timeZone: timezone },
    isAllDay: event.allDay,
    location: policy.includeLocation && event.location ? { displayName: event.location } : undefined,
    showAs: rule.transparency === "transparent" ? "free" : "busy",
    sensitivity: rule.visibility === "private" ? "private" : rule.visibility === "public" ? "normal" : "normal",
    isReminderOn: rule.reminderMode === "calendar_default" ? undefined : rule.reminderMode === "custom",
    reminderMinutesBeforeStart: rule.reminderMode === "custom" ? rule.reminderMinutes : undefined,
    transactionId: uuidFromHex(transactionHash),
  };
  return body;
}

export async function eventBody(
  event: NormalizedSchoolboxCalendarEvent,
  googleUserId: string,
  timezone: string,
  sourceKey: string,
  policy: SyncPolicy,
): Promise<GoogleCalendarEventInput & { id: string }> {
  const id = await createDeterministicEventId(`${googleUserId}:${sourceKey}`);
  const sourceLink = event.sourceUrl;
  const googleRule = resolveGoogleEventRule({ category: event.category ?? "other", type: event.type }, policy);
  const descriptionParts = [
    policy.includeDescription ? event.description : undefined,
    policy.includeEventTypeInDescription && event.type ? `Schoolbox type: ${event.type}` : undefined,
    policy.includeAuthorInDescription && event.author ? `Schoolbox author: ${event.author}` : undefined,
    policy.includeSchoolboxLink && sourceLink ? `Schoolbox: ${sourceLink}` : undefined,
  ].filter(Boolean);
  const body: GoogleCalendarEventInput & { id: string } = {
    id,
    summary: `${policy.titlePrefix ? `${policy.titlePrefix} ` : ""}${event.title || "Schoolbox event"}`,
    description: descriptionParts.join("\n\n") || undefined,
    location: policy.includeLocation ? event.location || undefined : undefined,
    start: event.allDay ? { date: event.start.slice(0, 10) } : { dateTime: event.start, timeZone: timezone },
    end: event.allDay ? { date: event.end.slice(0, 10) } : { dateTime: event.end, timeZone: timezone },
    extendedProperties: {
      private: {
        relaySource: "schoolbox",
        relaySourceKey: sourceKey.slice(0, 1024),
        relaySourceType: (event.type || "event").slice(0, 1024),
        relayManaged: "true",
      },
    },
  };
  if (policy.includeSchoolboxLink && sourceLink) body.source = { title: "Open in Schoolbox", url: sourceLink };
  if (googleRule.visibility !== "default") body.visibility = googleRule.visibility;
  if (googleRule.transparency !== "opaque") body.transparency = googleRule.transparency;
  if (googleRule.colorId) body.colorId = googleRule.colorId;
  if (googleRule.reminderMode === "none") body.reminders = { useDefault: false };
  if (googleRule.reminderMode === "custom") {
    body.reminders = {
      useDefault: false,
      overrides: [{ method: googleRule.reminderMethod, minutes: googleRule.reminderMinutes }],
    };
  }
  return body;
}

async function processInPool<T>(
  items: T[],
  concurrency: number,
  worker: (item: T) => Promise<void>,
  signal?: AbortSignal,
): Promise<void> {
  let index = 0;
  let failed = false;
  let failure: unknown;
  const runners = Array.from({ length: Math.min(Math.max(1, concurrency), Math.max(1, items.length)) }, async () => {
    while (index < items.length && !failed) {
      const item = items[index];
      index += 1;
      try {
        throwIfAborted(signal);
        await worker(item);
      } catch (error) {
        if (!failed) failure = error;
        failed = true;
      }
    }
  });
  await Promise.all(runners);
  if (failed) throw failure;
}

// Policy changes can require hundreds of independent Google event operations
// for one user. Keep enough parallelism to complete real reconciliation work
// inside the user deadline without allowing an unbounded API burst.
const USER_EVENT_OPERATION_CONCURRENCY = 6;
const MICROSOFT_EVENT_OPERATION_CONCURRENCY = 4;
// Increment this when Relay's Microsoft mutation semantics change. Including
// the revision in the stored source hash gives already-managed events one
// reconciliation pass instead of leaving values omitted by an older PATCH
// (and therefore preserved by Graph) in place indefinitely.
const MICROSOFT_EVENT_RECONCILIATION_VERSION = 2;

function microsoftEventPatchBody(body: MicrosoftEventInput): MicrosoftEventInput {
  const patch = { ...body };
  delete patch.transactionId;
  if (body.location) patch.location = body.location;
  else {
    delete patch.location;
    patch.locations = body.locations ?? [];
  }
  return patch;
}

async function microsoftReplacementEventBody(
  body: MicrosoftEventInput,
  replacedEventId: string,
): Promise<MicrosoftEventInput> {
  const replacementHash = await createContentHash({
    provider: "microsoft",
    transactionId: body.transactionId ?? null,
    replaces: replacedEventId,
  });
  return { ...body, transactionId: uuidFromHex(replacementHash) };
}

async function resolveMicrosoftCalendarId(options: {
  destinationId: string;
  userId: string;
  timezone: string;
  policy: SyncPolicy;
  microsoft: MicrosoftSyncClient;
  signal?: AbortSignal;
}): Promise<string> {
  throwIfAborted(options.signal);
  if (options.destinationId === "primary") return "primary";
  const definition = options.policy.secondaryCalendars.find((calendar) => calendar.id === options.destinationId);
  if (!definition) throw new Error(`Calendar destination ${options.destinationId} is no longer configured.`);
  const existing = await getUserCalendarTarget(options.userId, definition.id, "microsoft");
  const now = new Date().toISOString();
  if (existing) {
    if (existing.summary !== definition.name || existing.description !== definition.description) {
      await options.microsoft.updateCalendar(options.userId, existing.targetCalendarId, { name: definition.name }, { signal: options.signal });
      await upsertUserCalendarTarget({ ...existing, summary: definition.name, description: definition.description, updatedAt: now });
    }
    return existing.targetCalendarId;
  }
  const created = await options.microsoft.createCalendar(options.userId, { name: definition.name }, { signal: options.signal });
  const calendarId = created.id?.trim();
  if (!calendarId) throw new Error("Microsoft created a secondary calendar without returning its identifier.");
  await upsertUserCalendarTarget({
    target: "microsoft",
    targetUserId: options.userId,
    targetCalendarId: calendarId,
    googleUserId: options.userId,
    googleCalendarId: calendarId,
    destinationId: definition.id,
    summary: definition.name,
    description: definition.description,
    timeZone: options.timezone,
    createdAt: now,
    updatedAt: now,
  });
  return calendarId;
}

async function reconcileMicrosoftCalendars(options: {
  targets: UserCalendarTarget[];
  userId: string;
  policy: SyncPolicy;
  microsoft: MicrosoftSyncClient;
  signal?: AbortSignal;
}): Promise<void> {
  const configured = new Map(options.policy.secondaryCalendars.map((calendar) => [calendar.id, calendar]));
  for (const target of options.targets) {
    const definition = configured.get(target.destinationId);
    if (!definition || (target.summary === definition.name && target.description === definition.description)) continue;
    throwIfAborted(options.signal);
    await options.microsoft.updateCalendar(options.userId, target.targetCalendarId, { name: definition.name }, { signal: options.signal });
    await upsertUserCalendarTarget({ ...target, summary: definition.name, description: definition.description, updatedAt: new Date().toISOString() });
  }
}

async function syncMicrosoftUser(
  match: MicrosoftMatchedUser,
  run: RunSummary,
  targetRun: RunTargetSummary,
  schoolbox: SchoolboxSyncClient,
  microsoft: MicrosoftSyncClient,
  options: { pastDays: number; futureDays: number; timezone: string; syncPolicy: SyncPolicy; signal?: AbortSignal },
): Promise<void> {
  const target: TargetProvider = "microsoft";
  const userId = match.microsoft.id;
  const email = microsoftEmail(match.microsoft);
  const now = new Date();
  const windowStart = new Date(now.getTime() - options.pastDays * 86_400_000);
  const windowEnd = new Date(now.getTime() + options.futureDays * 86_400_000);
  const baseMapping = {
    target,
    targetUserId: userId,
    targetEmail: email,
    schoolboxUserId: match.schoolbox.id,
    schoolboxEmail: match.schoolboxEmail,
    displayName: microsoftDisplayName(match.microsoft) || schoolboxDisplayName(match.schoolbox),
    role: schoolboxRole(match.schoolbox),
    updatedAt: new Date().toISOString(),
  };
  const counters = { eventsFound: 0, eventsIncluded: 0, eventsCreated: 0, eventsUpdated: 0, eventsDeleted: 0, eventsUnchanged: 0, managedEventsAfter: 0 };
  let stage = "fetching_events";
  await startRunUserDiagnostic({ runId: run.id, target, targetUserId: userId, targetEmail: email,
    displayName: baseMapping.displayName, schoolboxUserId: match.schoolbox.id, schoolboxEmail: match.schoolboxEmail });

  const recordEvent = (event: NormalizedSchoolboxCalendarEvent, sourceKey: string, action: string, diagnostic: {
    detail?: string | null; errorMessage?: string | null; targetEventId?: string | null;
    calendarId?: string | null; destinationId?: string | null;
  } = {}) => recordRunEventDiagnostic({
    runId: run.id, target, targetUserId: userId, sourceKey, title: event.title,
    description: event.description, location: event.location, author: event.author,
    eventType: event.type, category: event.category ?? "other", sourceStart: event.start,
    sourceEnd: event.end, allDay: event.allDay, sourceUrl: event.sourceUrl,
    targetEventId: diagnostic.targetEventId ?? null, calendarId: diagnostic.calendarId ?? null,
    destinationId: diagnostic.destinationId ?? null, action, detail: diagnostic.detail ?? null,
    errorMessage: diagnostic.errorMessage ?? null,
  });

  try {
    const [events, storedMappings, storedTargets, exclusions] = await Promise.all([
      schoolbox.getCalendarEvents(match.schoolbox.id, { pastDays: options.pastDays, futureDays: options.futureDays, now, signal: options.signal }),
      getEventMappings(userId, target), listUserCalendarTargets(userId, target), getUserEventExclusions(userId, target),
    ]);
    counters.eventsFound = events.length;
    stage = "reconciling_calendars";
    await reconcileMicrosoftCalendars({ targets: storedTargets, userId, policy: options.syncPolicy, microsoft, signal: options.signal });
    const existing = new Map(storedMappings.map((mapping) => [mapping.sourceKey, mapping]));
    const seen = new Set<string>();
    const excluded = new Set<string>();
    const excludedRoots = new Set<string>();
    const userExcluded = new Set<string>();
    const userExcludedRoots = new Set<string>();
    const targetPromises = new Map<string, Promise<string>>();
    await recordDiscoveredEventTypes(events);
    stage = "processing_events";
    const targetFor = (destinationId: string) => {
      const cached = targetPromises.get(destinationId);
      if (cached) return cached;
      const pending = resolveMicrosoftCalendarId({ destinationId, userId, timezone: options.timezone, policy: options.syncPolicy, microsoft, signal: options.signal });
      targetPromises.set(destinationId, pending);
      return pending;
    };
    const included: Array<{ event: NormalizedSchoolboxCalendarEvent; sourceKey: string }> = [];
    for (const event of events) {
      throwIfAborted(options.signal);
      const sourceKey = `${event.sourceKey}:occurrence:${event.start}`;
      if (seen.has(sourceKey) || excluded.has(sourceKey)) continue;
      const policyEvent = { category: event.category ?? "other", type: event.type, allDay: event.allDay, completed: Boolean(event.completed) };
      const globallyIncluded = eventIncludedByPolicy(policyEvent, options.syncPolicy);
      const excludedForUser = globallyIncluded && eventExcludedForUser(policyEvent, exclusions);
      if (!globallyIncluded || excludedForUser) {
        excluded.add(sourceKey); excludedRoots.add(event.sourceKey);
        if (excludedForUser) { userExcluded.add(sourceKey); userExcludedRoots.add(event.sourceKey); }
        await recordEvent(event, sourceKey, "excluded", { detail: excludedForUser ? "Excluded by this person's custom settings." : "Excluded by Microsoft 365 target policy." });
        continue;
      }
      seen.add(sourceKey);
      included.push({ event, sourceKey });
    }

    await processInPool(included, MICROSOFT_EVENT_OPERATION_CONCURRENCY, async ({ event, sourceKey }) => {
      counters.eventsIncluded += 1;
      let calendarId: string | null = null;
      let destinationId: string | null = null;
      try {
        const rule = resolveGoogleEventRule({ category: event.category ?? "other", type: event.type }, options.syncPolicy);
        destinationId = rule.destinationId;
        calendarId = await targetFor(destinationId);
        const body = await microsoftEventBody(event, userId, options.timezone, sourceKey, options.syncPolicy, destinationId);
        const hash = await createContentHash({
          microsoftEventReconciliationVersion: MICROSOFT_EVENT_RECONCILIATION_VERSION,
          body,
        });
        const mapping = existing.get(sourceKey);
        const calendarChanged = Boolean(mapping && mapping.calendarId !== calendarId);
        if (mapping?.sourceHash === hash && !calendarChanged) {
          await touchEventMapping(userId, sourceKey, run.id, { title: event.title, description: event.description,
            location: event.location, author: event.author, eventType: event.type, category: event.category ?? "other",
            allDay: event.allDay, sourceUrl: event.sourceUrl, destinationId }, target);
          counters.eventsUnchanged += 1; targetRun.eventsUnchanged += 1;
          await recordEvent(event, sourceKey, "unchanged", { targetEventId: mapping.targetEventId, calendarId, destinationId, detail: "The managed Microsoft 365 event already matched the Schoolbox source." });
          return;
        }
        let eventId = mapping?.targetEventId;
        let createdAt = mapping?.createdAt ?? new Date().toISOString();
        let action: "created" | "updated" = "created";
        if (mapping && !calendarChanged && body.isReminderOn !== undefined) {
          try {
            const updated = await microsoft.updateEvent(
              userId,
              mapping.targetEventId,
              microsoftEventPatchBody(body),
              { calendarId: calendarId === "primary" ? undefined : calendarId, signal: options.signal },
            );
            eventId = updated.id?.trim() || mapping.targetEventId;
            counters.eventsUpdated += 1; targetRun.eventsUpdated += 1; action = "updated";
          } catch (error) {
            if (!(error instanceof MicrosoftGraphError) || (error.status !== 404 && error.status !== 410)) throw error;
            const created = await microsoft.insertEvent(
              userId,
              await microsoftReplacementEventBody(body, mapping.targetEventId),
              { calendarId: calendarId === "primary" ? undefined : calendarId, signal: options.signal },
            );
            eventId = created.id?.trim();
            if (!eventId) throw new Error("Microsoft created a replacement event without returning its identifier.");
            counters.eventsUpdated += 1; targetRun.eventsUpdated += 1; action = "updated";
            createdAt = new Date().toISOString();
          }
        } else {
          // Graph PATCH preserves omitted reminder properties. Recreating a
          // mapped event is therefore the only supported way to reapply the
          // mailbox's Outlook defaults without guessing a reminder interval.
          const createBody = mapping
            ? await microsoftReplacementEventBody(body, mapping.targetEventId)
            : body;
          const created = await microsoft.insertEvent(userId, createBody, { calendarId: calendarId === "primary" ? undefined : calendarId, signal: options.signal });
          eventId = created.id?.trim();
          if (!eventId) throw new Error("Microsoft created an event without returning its identifier.");
          if (mapping) {
            try { await microsoft.deleteEvent(userId, mapping.targetEventId, { calendarId: mapping.calendarId === "primary" ? undefined : mapping.calendarId, signal: options.signal }); }
            catch (error) { if (!(error instanceof MicrosoftGraphError) || (error.status !== 404 && error.status !== 410)) throw error; }
            counters.eventsUpdated += 1; targetRun.eventsUpdated += 1; action = "updated";
          } else { counters.eventsCreated += 1; targetRun.eventsCreated += 1; }
          createdAt = new Date().toISOString();
        }
        await upsertEventMapping({ target, targetUserId: userId, targetEventId: eventId, sourceKey,
          calendarId, sourceHash: hash, sourceStart: event.start, sourceEnd: event.end,
          lastSeenRunId: run.id, createdAt, updatedAt: new Date().toISOString(), title: event.title,
          description: event.description, location: event.location, author: event.author,
          eventType: event.type, category: event.category ?? "other", allDay: event.allDay,
          sourceUrl: event.sourceUrl, destinationId });
        await recordEvent(event, sourceKey, action, { targetEventId: eventId, calendarId, destinationId,
          detail: calendarChanged ? "The managed event was moved to its configured Microsoft 365 calendar." : `A managed Microsoft 365 event was ${action}.` });
      } catch (error) {
        await recordEvent(event, sourceKey, "failed", { calendarId, destinationId, detail: "Failure while processing this Microsoft 365 event.", errorMessage: syncErrorMessage(error, options.signal) });
        throw error;
      }
    }, options.signal);

    stage = "reconciling_removed_events";
    const removals = storedMappings.filter((mapping) => {
      if (seen.has(mapping.sourceKey)) return false;
      const marker = mapping.sourceKey.lastIndexOf(":occurrence:");
      const root = marker >= 0 ? mapping.sourceKey.slice(0, marker) : mapping.sourceKey;
      const excludedByPolicy = excluded.has(mapping.sourceKey) || excludedRoots.has(root);
      if (excludedByPolicy ? !options.syncPolicy.deleteExcludedEvents : !options.syncPolicy.deleteMissingEvents) return false;
      const sourceStart = new Date(mapping.sourceStart); const sourceEnd = new Date(mapping.sourceEnd);
      return sourceStart < windowEnd && sourceEnd > windowStart;
    });
    await processInPool(removals, MICROSOFT_EVENT_OPERATION_CONCURRENCY, async (mapping) => {
      try { await microsoft.deleteEvent(userId, mapping.targetEventId, { calendarId: mapping.calendarId === "primary" ? undefined : mapping.calendarId, signal: options.signal }); }
      catch (error) { if (!(error instanceof MicrosoftGraphError) || (error.status !== 404 && error.status !== 410)) throw error; }
      await deleteEventMapping(userId, mapping.sourceKey, target);
      counters.eventsDeleted += 1; targetRun.eventsDeleted += 1;
      const marker = mapping.sourceKey.lastIndexOf(":occurrence:");
      const root = marker >= 0 ? mapping.sourceKey.slice(0, marker) : mapping.sourceKey;
      await recordRunEventDiagnostic({ runId: run.id, target, targetUserId: userId, sourceKey: mapping.sourceKey,
        title: mapping.title, description: mapping.description, location: mapping.location, author: mapping.author,
        eventType: mapping.eventType, category: mapping.category, sourceStart: mapping.sourceStart,
        sourceEnd: mapping.sourceEnd, allDay: mapping.allDay, sourceUrl: mapping.sourceUrl,
        targetEventId: mapping.targetEventId, calendarId: mapping.calendarId, destinationId: mapping.destinationId,
        action: "deleted", detail: userExcluded.has(mapping.sourceKey) || userExcludedRoots.has(root)
          ? "The managed Microsoft 365 event was deleted because it is excluded for this person."
          : "The managed Microsoft 365 event was deleted during reconciliation.", errorMessage: null });
    }, options.signal);
    stage = "saving_user_result";
    counters.managedEventsAfter = (await getEventMappings(userId, target)).length;
    await upsertUserMapping({ ...baseMapping, status: "synced", lastSyncAt: new Date().toISOString(), lastError: null, eventCount: counters.managedEventsAfter });
    await finishRunUserDiagnostic({ runId: run.id, target, targetUserId: userId, status: "completed", stage: "completed", ...counters, errorMessage: null });
    targetRun.usersSynced += 1;
  } catch (error) {
    const message = syncErrorMessage(error, options.signal);
    try { counters.managedEventsAfter = (await getEventMappings(userId, target)).length; } catch { /* preserve original error */ }
    await upsertUserMapping({ ...baseMapping, status: "error", lastSyncAt: new Date().toISOString(), lastError: message.slice(0, 2_000), eventCount: counters.managedEventsAfter });
    await finishRunUserDiagnostic({ runId: run.id, target, targetUserId: userId, status: "failed", stage, ...counters, errorMessage: message });
    targetRun.errors += 1;
  }
}

async function resolveCalendarId(options: {
  destinationId: string;
  googleUserId: string;
  googleEmail: string;
  timezone: string;
  policy: SyncPolicy;
  google: GoogleSyncClient;
  signal?: AbortSignal;
}): Promise<string> {
  throwIfAborted(options.signal);
  if (options.destinationId === "primary") return "primary";
  const definition = options.policy.secondaryCalendars.find((calendar) => calendar.id === options.destinationId);
  if (!definition) throw new Error(`Calendar destination ${options.destinationId} is no longer configured.`);

  const existing = await getUserCalendarTarget(options.googleUserId, definition.id);
  const expected = {
    summary: definition.name,
    description: definition.description,
    timeZone: options.timezone,
  };
  const now = new Date().toISOString();
  if (existing) {
    if (
      existing.summary !== expected.summary ||
      existing.description !== expected.description ||
      existing.timeZone !== expected.timeZone
    ) {
      await options.google.updateCalendar(options.googleEmail, existing.googleCalendarId, expected, {
        quotaUser: options.googleUserId,
        signal: options.signal,
      });
      await upsertUserCalendarTarget({
        ...existing,
        ...expected,
        updatedAt: now,
      });
    }
    return existing.googleCalendarId;
  }

  const created = await options.google.createCalendar(options.googleEmail, expected, {
    quotaUser: options.googleUserId,
    signal: options.signal,
  });
  const googleCalendarId = created.id?.trim();
  if (!googleCalendarId) throw new Error("Google created a secondary calendar without returning its identifier.");
  await upsertUserCalendarTarget({
    googleUserId: options.googleUserId,
    destinationId: definition.id,
    googleCalendarId,
    ...expected,
    createdAt: now,
    updatedAt: now,
  });
  return googleCalendarId;
}

async function reconcileExistingCalendarTargets(options: {
  targets: UserCalendarTarget[];
  googleUserId: string;
  googleEmail: string;
  timezone: string;
  policy: SyncPolicy;
  google: GoogleSyncClient;
  signal?: AbortSignal;
}): Promise<void> {
  const configured = new Map(options.policy.secondaryCalendars.map((calendar) => [calendar.id, calendar]));
  for (const target of options.targets) {
    throwIfAborted(options.signal);
    const definition = configured.get(target.destinationId);
    if (!definition) continue;
    const expected = {
      summary: definition.name,
      description: definition.description,
      timeZone: options.timezone,
    };
    if (
      target.summary === expected.summary &&
      target.description === expected.description &&
      target.timeZone === expected.timeZone
    ) continue;

    await options.google.updateCalendar(options.googleEmail, target.googleCalendarId, expected, {
      quotaUser: options.googleUserId,
      signal: options.signal,
    });
    await upsertUserCalendarTarget({
      ...target,
      ...expected,
      updatedAt: new Date().toISOString(),
    });
  }
}

async function syncUser(
  match: MatchedUser,
  run: RunSummary,
  schoolbox: SchoolboxSyncClient,
  google: GoogleSyncClient,
  options: { pastDays: number; futureDays: number; timezone: string; syncPolicy: SyncPolicy; signal?: AbortSignal },
): Promise<void> {
  const now = new Date();
  const windowStart = new Date(now.getTime() - options.pastDays * 86_400_000);
  const windowEnd = new Date(now.getTime() + options.futureDays * 86_400_000);
  const googleUserId = match.google.id;
  const googleEmail = match.google.primaryEmail;
  const baseMapping = {
    googleUserId,
    googleEmail,
    schoolboxUserId: match.schoolbox.id,
    schoolboxEmail: match.schoolboxEmail,
    displayName: googleDisplayName(match.google) || schoolboxDisplayName(match.schoolbox),
    role: schoolboxRole(match.schoolbox),
    updatedAt: new Date().toISOString(),
  };
  const counters = {
    eventsFound: 0,
    eventsIncluded: 0,
    eventsCreated: 0,
    eventsUpdated: 0,
    eventsDeleted: 0,
    eventsUnchanged: 0,
    managedEventsAfter: 0,
  };
  let stage = "fetching_events";

  await startRunUserDiagnostic({
    runId: run.id,
    googleUserId,
    googleEmail,
    displayName: baseMapping.displayName,
    schoolboxUserId: match.schoolbox.id,
    schoolboxEmail: match.schoolboxEmail,
  });

  const recordEvent = async (
    event: NormalizedSchoolboxCalendarEvent,
    sourceKey: string,
    action: string,
    diagnostic: {
      detail?: string | null;
      errorMessage?: string | null;
      googleEventId?: string | null;
      calendarId?: string | null;
      destinationId?: string | null;
    } = {},
  ) => recordRunEventDiagnostic({
    runId: run.id,
    googleUserId,
    sourceKey,
    title: event.title,
    description: event.description,
    location: event.location,
    author: event.author,
    eventType: event.type,
    category: event.category ?? "other",
    sourceStart: event.start,
    sourceEnd: event.end,
    allDay: event.allDay,
    sourceUrl: event.sourceUrl,
    googleEventId: diagnostic.googleEventId ?? null,
    calendarId: diagnostic.calendarId ?? null,
    destinationId: diagnostic.destinationId ?? null,
    action,
    detail: diagnostic.detail ?? null,
    errorMessage: diagnostic.errorMessage ?? null,
  });

  try {
    const [events, storedMappings, storedCalendarTargets, userExclusions] = await Promise.all([
      schoolbox.getCalendarEvents(match.schoolbox.id, {
        pastDays: options.pastDays,
        futureDays: options.futureDays,
        now,
        signal: options.signal,
      }),
      getEventMappings(googleUserId),
      listUserCalendarTargets(googleUserId),
      getUserEventExclusions(googleUserId),
    ]);
    counters.eventsFound = events.length;
    stage = "reconciling_calendars";
    await reconcileExistingCalendarTargets({
      targets: storedCalendarTargets,
      googleUserId,
      googleEmail,
      timezone: options.timezone,
      policy: options.syncPolicy,
      google,
      signal: options.signal,
    });
    const existing = new Map(storedMappings.map((mapping) => [mapping.sourceKey, mapping]));
    const seen = new Set<string>();
    const excluded = new Set<string>();
    const excludedSourceRoots = new Set<string>();
    const userExcluded = new Set<string>();
    const userExcludedSourceRoots = new Set<string>();
    const calendarTargets = new Map<string, Promise<string>>();
    await recordDiscoveredEventTypes(events);
    stage = "processing_events";

    const targetFor = (destinationId: string) => {
      const cached = calendarTargets.get(destinationId);
      if (cached) return cached;
      const target = resolveCalendarId({
        destinationId,
        googleUserId,
        googleEmail,
        timezone: options.timezone,
        policy: options.syncPolicy,
        google,
        signal: options.signal,
      });
      calendarTargets.set(destinationId, target);
      return target;
    };

    const includedEvents: Array<{ event: NormalizedSchoolboxCalendarEvent; sourceKey: string }> = [];
    for (const event of events) {
      throwIfAborted(options.signal);
      const sourceKey = `${event.sourceKey}:occurrence:${event.start}`;
      if (seen.has(sourceKey) || excluded.has(sourceKey)) continue;
      const policyEvent = {
        category: event.category ?? "other",
        type: event.type,
        allDay: event.allDay,
        completed: Boolean(event.completed),
      };
      const globallyIncluded = eventIncludedByPolicy(policyEvent, options.syncPolicy);
      const excludedForUser = globallyIncluded && eventExcludedForUser(policyEvent, userExclusions);
      if (!globallyIncluded || excludedForUser) {
        excluded.add(sourceKey);
        excludedSourceRoots.add(event.sourceKey);
        if (excludedForUser) {
          userExcluded.add(sourceKey);
          userExcludedSourceRoots.add(event.sourceKey);
        }
        await recordEvent(event, sourceKey, "excluded", {
          detail: options.syncPolicy.deleteExcludedEvents
            ? excludedForUser
              ? "Excluded by this person's custom settings; any tracked Google copy will be removed."
              : "Excluded by organisation policy; any tracked Google copy will be removed."
            : excludedForUser
              ? "Excluded by this person's custom settings; an existing Google copy is retained by configuration."
              : "Excluded by organisation policy; an existing Google copy is retained by configuration.",
        });
        continue;
      }
      seen.add(sourceKey);
      includedEvents.push({ event, sourceKey });
    }

    await processInPool(includedEvents, USER_EVENT_OPERATION_CONCURRENCY, async ({ event, sourceKey }) => {
      counters.eventsIncluded += 1;
      let calendarId: string | null = null;
      let destinationId: string | null = null;
      try {
        const googleRule = resolveGoogleEventRule({ category: event.category ?? "other", type: event.type }, options.syncPolicy);
        destinationId = googleRule.destinationId;
        calendarId = await targetFor(googleRule.destinationId);
        const body = await eventBody(event, googleUserId, options.timezone, sourceKey, options.syncPolicy);
        const hash = await createContentHash(body);
        const mapping = existing.get(sourceKey);
        const calendarChanged = Boolean(mapping && mapping.calendarId !== calendarId);

        if (mapping?.sourceHash === hash && !calendarChanged) {
          await touchEventMapping(googleUserId, sourceKey, run.id, {
            title: event.title,
            description: event.description,
            location: event.location,
            author: event.author,
            eventType: event.type,
            category: event.category ?? "other",
            allDay: event.allDay,
            sourceUrl: event.sourceUrl,
            destinationId: googleRule.destinationId,
          });
          run.eventsUnchanged += 1;
          counters.eventsUnchanged += 1;
          await recordEvent(event, sourceKey, "unchanged", {
            googleEventId: mapping.googleEventId,
            calendarId,
            destinationId: googleRule.destinationId,
            detail: "The managed Google event already matched the Schoolbox source.",
          });
          return;
        }

        let createdAt = mapping?.createdAt ?? new Date().toISOString();
        let action: "created" | "updated" = "created";
        if (mapping && !calendarChanged) {
          await google.updateEvent(googleEmail, mapping.googleEventId, body, {
            calendarId,
            quotaUser: googleUserId,
            sendUpdates: "none",
            signal: options.signal,
          });
          run.eventsUpdated += 1;
          counters.eventsUpdated += 1;
          action = "updated";
        } else {
          let insertConflict = false;
          try {
            await google.insertEvent(googleEmail, body, {
              calendarId,
              quotaUser: googleUserId,
              sendUpdates: "none",
              signal: options.signal,
            });
          } catch (error) {
            if (!(error instanceof GoogleApiError) || error.status !== 409) throw error;
            insertConflict = true;
            await google.updateEvent(googleEmail, body.id, body, {
              calendarId,
              quotaUser: googleUserId,
              sendUpdates: "none",
              signal: options.signal,
            });
            createdAt = new Date().toISOString();
          }
          if (mapping) {
            try {
              await google.deleteEvent(googleEmail, mapping.googleEventId, {
                calendarId: mapping.calendarId,
                quotaUser: googleUserId,
                sendUpdates: "none",
                signal: options.signal,
              });
            } catch (error) {
              if (!(error instanceof GoogleApiError) || (error.status !== 404 && error.status !== 410)) throw error;
            }
            run.eventsUpdated += 1;
            counters.eventsUpdated += 1;
            action = "updated";
          } else if (insertConflict) {
            run.eventsUpdated += 1;
            counters.eventsUpdated += 1;
            action = "updated";
          } else {
            run.eventsCreated += 1;
            counters.eventsCreated += 1;
          }
        }

        await upsertEventMapping({
          googleUserId,
          sourceKey,
          googleEventId: body.id,
          calendarId,
          sourceHash: hash,
          sourceStart: event.start,
          sourceEnd: event.end,
          lastSeenRunId: run.id,
          createdAt,
          updatedAt: new Date().toISOString(),
          title: event.title,
          description: event.description,
          location: event.location,
          author: event.author,
          eventType: event.type,
          category: event.category ?? "other",
          allDay: event.allDay,
          sourceUrl: event.sourceUrl,
          destinationId: googleRule.destinationId,
        });
        await recordEvent(event, sourceKey, action, {
          googleEventId: body.id,
          calendarId,
          destinationId: googleRule.destinationId,
          detail: calendarChanged
            ? "The managed event was moved to its configured calendar destination."
            : action === "created" ? "A new managed Google event was created." : "The managed Google event was updated.",
        });
      } catch (error) {
        const message = syncErrorMessage(error, options.signal);
        await recordEvent(event, sourceKey, "failed", {
          googleEventId: null,
          calendarId,
          destinationId,
          detail: "Failure while processing this managed event.",
          errorMessage: message,
        });
        throw error;
      }
    }, options.signal);

    stage = "reconciling_removed_events";
    const removalCandidates: Array<{
      mapping: EventMapping;
      excludedByPolicy: boolean;
      excludedByUser: boolean;
    }> = [];
    for (const mapping of storedMappings) {
      throwIfAborted(options.signal);
      if (seen.has(mapping.sourceKey)) continue;
      const occurrenceMarker = mapping.sourceKey.lastIndexOf(":occurrence:");
      const mappingRoot = occurrenceMarker >= 0 ? mapping.sourceKey.slice(0, occurrenceMarker) : mapping.sourceKey;
      const excludedByPolicy = excluded.has(mapping.sourceKey) || excludedSourceRoots.has(mappingRoot);
      const excludedByUser = userExcluded.has(mapping.sourceKey) || userExcludedSourceRoots.has(mappingRoot);
      if (excludedByPolicy ? !options.syncPolicy.deleteExcludedEvents : !options.syncPolicy.deleteMissingEvents) continue;
      const sourceStart = new Date(mapping.sourceStart);
      const sourceEnd = new Date(mapping.sourceEnd);
      const wasInsideFetchedWindow = sourceStart < windowEnd && sourceEnd > windowStart;
      if (!wasInsideFetchedWindow) continue;
      removalCandidates.push({ mapping, excludedByPolicy, excludedByUser });
    }
    await processInPool(removalCandidates, USER_EVENT_OPERATION_CONCURRENCY, async ({
      mapping,
      excludedByPolicy,
      excludedByUser,
    }) => {
      try {
        await google.deleteEvent(googleEmail, mapping.googleEventId, {
          calendarId: mapping.calendarId,
          quotaUser: googleUserId,
          sendUpdates: "none",
          signal: options.signal,
        });
      } catch (error) {
        if (!(error instanceof GoogleApiError) || (error.status !== 404 && error.status !== 410)) {
          const message = syncErrorMessage(error, options.signal);
          await recordRunEventDiagnostic({
            runId: run.id,
            googleUserId,
            sourceKey: mapping.sourceKey,
            title: mapping.title,
            description: mapping.description,
            location: mapping.location,
            author: mapping.author,
            eventType: mapping.eventType,
            category: mapping.category,
            sourceStart: mapping.sourceStart,
            sourceEnd: mapping.sourceEnd,
            allDay: mapping.allDay,
            sourceUrl: mapping.sourceUrl,
            googleEventId: mapping.googleEventId,
            calendarId: mapping.calendarId,
            destinationId: mapping.destinationId,
            action: "failed",
            detail: "Failure while removing an excluded or missing managed event.",
            errorMessage: message,
          });
          throw error;
        }
      }
      await deleteEventMapping(googleUserId, mapping.sourceKey);
      run.eventsDeleted += 1;
      counters.eventsDeleted += 1;
      await recordRunEventDiagnostic({
        runId: run.id,
        googleUserId,
        sourceKey: mapping.sourceKey,
        title: mapping.title,
        description: mapping.description,
        location: mapping.location,
        author: mapping.author,
        eventType: mapping.eventType,
        category: mapping.category,
        sourceStart: mapping.sourceStart,
        sourceEnd: mapping.sourceEnd,
        allDay: mapping.allDay,
        sourceUrl: mapping.sourceUrl,
        googleEventId: mapping.googleEventId,
        calendarId: mapping.calendarId,
        destinationId: mapping.destinationId,
        action: "deleted",
        detail: excludedByUser
          ? "The managed Google event was deleted because it is excluded for this person."
          : excludedByPolicy
            ? "The managed Google event was deleted because the source is excluded by organisation policy."
          : "The managed Google event was deleted because it was no longer returned by Schoolbox.",
        errorMessage: null,
      });
    }, options.signal);

    throwIfAborted(options.signal);
    stage = "saving_user_result";
    const managedEventCount = (await getEventMappings(googleUserId)).length;
    counters.managedEventsAfter = managedEventCount;
    await upsertUserMapping({
      ...baseMapping,
      status: "synced",
      lastSyncAt: new Date().toISOString(),
      lastError: null,
      eventCount: managedEventCount,
    });
    await finishRunUserDiagnostic({
      runId: run.id,
      googleUserId,
      status: "completed",
      stage: "completed",
      ...counters,
      errorMessage: null,
    });
    run.usersSynced += 1;
  } catch (error) {
    const message = syncErrorMessage(error, options.signal);
    let managedEventsOnFailure = 0;
    try {
      managedEventsOnFailure = (await getEventMappings(googleUserId)).length;
    } catch {
      // Preserve the original sync failure if the diagnostic count cannot be read.
    }
    counters.managedEventsAfter = managedEventsOnFailure;
    await upsertUserMapping({
      ...baseMapping,
      status: "error",
      lastSyncAt: new Date().toISOString(),
      lastError: message.slice(0, 2000),
      eventCount: managedEventsOnFailure,
    });
    await finishRunUserDiagnostic({
      runId: run.id,
      googleUserId,
      status: "failed",
      stage,
      ...counters,
      errorMessage: message,
    });
    run.errors += 1;
  }
}

/**
 * Pauses one target account and removes only events tracked in Relay's
 * provider-qualified mapping table. When explicitly requested, it then
 * permanently deletes only the tracked secondary calendars. Deleting a calendar
 * also deletes any manually added content it contains, so this path never
 * accepts a primary or untracked calendar ID.
 */
export async function cleanupUserManagedEvents(
  targetUserId: string,
  actor: string,
  clientOverride?: TargetCleanupClient,
  options: { deleteCalendars?: boolean; target?: TargetProvider; timeoutMs?: number } = {},
): Promise<ManagedEventCleanupResult> {
  const target = options.target ?? "google";
  const userId = targetUserId.trim();
  if (!userId) throw new HttpError(400, "Choose a user to clean up");

  const mapping = await getUserMapping(userId, target);
  if (!mapping) throw new HttpError(404, "This user is no longer available");

  const [storedMappings, storedCalendarTargets] = await Promise.all([
    getEventMappings(userId, target),
    listUserCalendarTargets(userId, target),
  ]);
  const deleteCalendars = options.deleteCalendars === true;
  const hasDestructiveWork = storedMappings.length > 0 || (deleteCalendars && storedCalendarTargets.length > 0);
  if (hasDestructiveWork) await recoverStaleRuns();
  if (hasDestructiveWork && (await listRuns(1))[0]?.status === "running") {
    throw new HttpError(
      409,
      "Wait for the current calendar operation to finish, then retry cleanup.",
    );
  }

  // Pausing before the first delete prevents future scheduled runs from
  // recreating the events immediately after a successful cleanup.
  await setUsersSyncEnabled([userId], false, actor, target);
  if (!hasDestructiveWork) {
    await recordManagedEventCleanup({
      target,
      targetUserId: userId,
      remaining: 0,
      deleted: 0,
      alreadyMissing: 0,
      calendarsRemaining: storedCalendarTargets.length,
      error: null,
      actor,
    });
    return {
      paused: true,
      deleted: 0,
      alreadyMissing: 0,
      remaining: 0,
      calendarsDeleted: 0,
      calendarsAlreadyMissing: 0,
      calendarsRemaining: storedCalendarTargets.length,
      error: null,
    };
  }

  let client: TargetCleanupClient;
  let cleanupTimeoutMs: number;
  try {
    const config = await getConfig(!clientOverride);
    cleanupTimeoutMs = positiveTimeout(options.timeoutMs, config.userSyncTimeoutSeconds * 1_000);
    if (clientOverride) {
      client = clientOverride;
    } else {
      if (target === "google") {
        if (!config.googleServiceAccountJson) throw new HttpError(409, "Google Workspace is not configured");
        client = new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson));
      } else {
        if (!config.microsoftTenantId || !config.microsoftClientId || !config.microsoftClientSecret) throw new HttpError(409, "Microsoft 365 is not configured");
        client = new MicrosoftGraphClient({ tenantId: config.microsoftTenantId, clientId: config.microsoftClientId, clientSecret: config.microsoftClientSecret });
      }
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Calendar target cleanup could not start";
    await recordManagedEventCleanup({
      target,
      targetUserId: userId,
      remaining: storedMappings.length,
      deleted: 0,
      alreadyMissing: 0,
      calendarsRemaining: storedCalendarTargets.length,
      error: message,
      actor,
    });
    throw error;
  }

  let deleted = 0;
  let alreadyMissing = 0;
  let cleanupError: string | null = null;
  let calendarsDeleted = 0;
  let calendarsAlreadyMissing = 0;
  try {
    await withHardDeadline(
      undefined,
      cleanupTimeoutMs,
      `Managed ${target === "google" ? "Google" : "Microsoft 365"} cleanup exceeded ${durationLabel(cleanupTimeoutMs)}`,
      async (signal) => {
        for (const eventMapping of storedMappings) {
          throwIfAborted(signal);
          let removalOutcome: "deleted" | "already_missing" | null = null;
          try {
            if (target === "google") await (client as GoogleCleanupClient).deleteEvent(mapping.targetEmail, eventMapping.targetEventId, {
              calendarId: eventMapping.calendarId, quotaUser: userId, sendUpdates: "none", signal,
            });
            else await (client as MicrosoftCleanupClient).deleteEvent(userId, eventMapping.targetEventId, {
              calendarId: eventMapping.calendarId === "primary" ? undefined : eventMapping.calendarId,
              signal,
            });
            removalOutcome = "deleted";
          } catch (error) {
            if ((error instanceof GoogleApiError || error instanceof MicrosoftGraphError) && (error.status === 404 || error.status === 410)) {
              removalOutcome = "already_missing";
            } else {
              cleanupError = syncErrorMessage(error, signal);
              break;
            }
          }
          if (removalOutcome) {
            await deleteEventMapping(userId, eventMapping.sourceKey, target);
            if (removalOutcome === "deleted") deleted += 1;
            else alreadyMissing += 1;
          }
        }

        const remainingEvents = (await getEventMappings(userId, target)).length;
        if (!deleteCalendars || remainingEvents > 0 || storedCalendarTargets.length === 0) return;
        if (!client.deleteCalendar) {
          cleanupError = `${target === "google" ? "Google" : "Microsoft 365"} cleanup does not support deleting secondary calendars`;
          return;
        }
        for (const calendarTarget of storedCalendarTargets) {
          throwIfAborted(signal);
          let removalOutcome: "deleted" | "already_missing" | null = null;
          try {
            if (target === "microsoft") await (client as MicrosoftCleanupClient).deleteCalendar!(userId, calendarTarget.targetCalendarId, { signal });
            else await (client as GoogleCleanupClient).deleteCalendar!(mapping.targetEmail, calendarTarget.targetCalendarId, { quotaUser: userId, signal });
            removalOutcome = "deleted";
          } catch (error) {
            if ((error instanceof GoogleApiError || error instanceof MicrosoftGraphError) && (error.status === 404 || error.status === 410)) {
              removalOutcome = "already_missing";
            } else {
              cleanupError = syncErrorMessage(error, signal);
            }
          }
          if (removalOutcome) {
            await deleteCalendarTargetRecords(userId, calendarTarget.destinationId, calendarTarget.targetCalendarId, target);
            if (removalOutcome === "deleted") calendarsDeleted += 1;
            else calendarsAlreadyMissing += 1;
          }
        }
      },
    );
  } catch (error) {
    cleanupError = syncErrorMessage(error);
  }

  const remaining = (await getEventMappings(userId, target)).length;
  const calendarsRemaining = (await listUserCalendarTargets(userId, target)).length;
  await recordManagedEventCleanup({
    target,
    targetUserId: userId,
    remaining,
    deleted,
    alreadyMissing,
    calendarsDeleted,
    calendarsAlreadyMissing,
    calendarsRemaining,
    error: cleanupError,
    actor,
  });
  return {
    paused: true,
    deleted,
    alreadyMissing,
    remaining,
    calendarsDeleted,
    calendarsAlreadyMissing,
    calendarsRemaining,
    error: cleanupError,
  };
}

/**
 * Permanently removes one managed secondary-calendar destination from every
 * tracked user. The destination is removed from policy before provider deletion
 * starts so a subsequent sync cannot recreate it. Failed target records are
 * retained and shown as retired cleanup work that an administrator can retry.
 */
export async function retireCalendarDestination(
  destinationId: string,
  actor: string,
  clientOverride?: CalendarRetirementClient,
  target: TargetProvider = "google",
  runtimeOptions: Pick<SyncRuntimeOptions, "runTimeoutMs"> = {},
): Promise<CalendarDestinationRetirementResult> {
  const id = destinationId.trim();
  if (!/^[a-z0-9][a-z0-9_-]{0,39}$/.test(id)) throw new HttpError(400, "Choose a valid calendar destination");
  if (id === "primary") throw new HttpError(400, "The primary calendar cannot be deleted");

  const initialConfig = await getConfig(false);
  const initialPolicy = target === "google" ? initialConfig.syncPolicy : initialConfig.microsoftSyncPolicy;
  const initialTargets = await listCalendarTargetsForDestination(id, target);
  if (!initialPolicy.secondaryCalendars.some((calendar) => calendar.id === id) && initialTargets.length === 0) {
    throw new HttpError(404, "This calendar destination is no longer available");
  }

  await recoverStaleRuns();
  const run = await createRun("calendar_retirement");
  const heartbeat = setInterval(() => {
    void touchRunHeartbeat(run.id).catch(() => undefined);
  }, 30_000);
  const result: CalendarDestinationRetirementResult = {
    destinationId: id,
    calendarsDeleted: 0,
    calendarsAlreadyMissing: 0,
    calendarsFailed: 0,
    calendarsRemaining: 0,
    eventMappingsRemoved: 0,
    error: null,
  };
  let unexpectedError: unknown;

  try {
    const config = await getConfig(true);
    const retirementTimeoutMs = positiveTimeout(runtimeOptions.runTimeoutMs, config.runTimeoutMinutes * 60_000);
    await withHardDeadline(
      undefined,
      retirementTimeoutMs,
      `Calendar destination retirement exceeded ${durationLabel(retirementTimeoutMs)}`,
      async (signal) => {
        const targets = await listCalendarTargetsForDestination(id, target);
        run.usersDiscovered = targets.length;
        run.usersMatched = targets.length;
        await checkpointRun(run, "calendar_retirement", `${targets.length} tracked user calendar(s) queued for retirement.`);

        let calendarClient: CalendarRetirementClient | null = clientOverride ?? null;
        if (targets.length > 0 && !calendarClient) {
          if (target === "google") {
            if (!config.googleServiceAccountJson) throw new HttpError(409, "Google Workspace is not configured");
            calendarClient = new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson));
          } else {
            if (!config.microsoftTenantId || !config.microsoftClientId || !config.microsoftClientSecret) {
              throw new HttpError(409, "Microsoft 365 is not configured");
            }
            calendarClient = new MicrosoftGraphClient({
              tenantId: config.microsoftTenantId,
              clientId: config.microsoftClientId,
              clientSecret: config.microsoftClientSecret,
            });
          }
        }

        const policy = target === "google" ? config.syncPolicy : config.microsoftSyncPolicy;
        if (policy.secondaryCalendars.some((calendar) => calendar.id === id)) {
          await saveConfig(target === "google"
            ? { syncPolicy: withoutManagedCalendarDestination(policy, id) }
            : { microsoftSyncPolicy: withoutManagedCalendarDestination(policy, id) }, actor);
        }

        const errors: string[] = [];
        await processInPool(targets, config.concurrency, async (calendarTarget) => {
          let removalOutcome: "deleted" | "already_missing" | null = null;
          try {
            if (calendarTarget.target === "google") {
              await (calendarClient as GoogleCalendarRetirementClient).deleteCalendar(
                calendarTarget.targetEmail,
                calendarTarget.targetCalendarId,
                { quotaUser: calendarTarget.targetUserId, signal },
              );
            } else {
              await (calendarClient as MicrosoftCalendarRetirementClient).deleteCalendar(
                calendarTarget.targetUserId,
                calendarTarget.targetCalendarId,
                { signal },
              );
            }
            removalOutcome = "deleted";
          } catch (error) {
            if ((error instanceof GoogleApiError || error instanceof MicrosoftGraphError) &&
                (error.status === 404 || error.status === 410)) {
              removalOutcome = "already_missing";
            } else {
              result.calendarsFailed += 1;
              const message = syncErrorMessage(error, signal);
              if (!errors.includes(message)) errors.push(message);
            }
          }
          if (removalOutcome) {
            const removed = await deleteCalendarTargetRecords(
              calendarTarget.targetUserId,
              id,
              calendarTarget.targetCalendarId,
              calendarTarget.target,
            );
            result.eventMappingsRemoved += removed;
            if (removalOutcome === "deleted") result.calendarsDeleted += 1;
            else result.calendarsAlreadyMissing += 1;
            run.eventsDeleted += removed;
            run.usersSynced += 1;
          }
          await checkpointRun(
            run,
            "calendar_retirement",
            `${run.usersSynced + result.calendarsFailed} of ${targets.length} tracked user calendar(s) processed.`,
          );
        }, signal);

        result.calendarsRemaining = (await listCalendarTargetsForDestination(id, target)).length;
        result.error = errors.length > 0 ? errors.join("; ").slice(0, 2_000) : null;
        run.errors = result.calendarsFailed;
        run.status = result.calendarsFailed > 0 ? "completed_with_errors" : "completed";
        run.message = result.calendarsFailed > 0
          ? `Calendar destination ${id} was retired, but ${result.calendarsRemaining} user calendar(s) still require deletion.`
          : `Calendar destination ${id} was retired and ${result.calendarsDeleted + result.calendarsAlreadyMissing} tracked user calendar(s) were removed.`;
      },
    );
  } catch (error) {
    unexpectedError = error;
    try {
      result.calendarsRemaining = (await listCalendarTargetsForDestination(id, target)).length;
    } catch {
      // Preserve the retirement failure when remaining-work diagnostics fail.
    }
    result.error = syncErrorMessage(error);
    run.status = "failed";
    run.errors += 1;
    run.message = result.error || "Calendar destination retirement failed.";
  } finally {
    clearInterval(heartbeat);
    run.completedAt = new Date().toISOString();
    await finishRun(run);
    await addAudit(actor, `calendar_destination.${run.status}`, `Run ${run.id}: ${run.message ?? run.status}`);
  }

  if (unexpectedError) throw unexpectedError;
  return result;
}

export async function runFullSync(
  trigger: string,
  actor: string,
  clientOverrides: SyncClientOverrides = {},
  runtimeOptions: SyncRuntimeOptions = {},
  requestedTargets?: TargetProvider[],
): Promise<RunSummary> {
  const config = await getConfig(true);
  if (!config.schoolboxBaseUrl || !config.schoolboxToken) {
    throw new HttpError(409, "Schoolbox is not configured");
  }
  if (!config.schoolboxSetupCompleted) {
    throw new HttpError(409, "Schoolbox setup has not been completed and verified");
  }
  const targetEnabled = (target: TargetProvider) => target === "google"
    ? config.googleEnabled
    : config.microsoftEnabled;
  const targetSetupCompleted = (target: TargetProvider) => target === "google"
    ? config.googleSetupCompleted
    : config.microsoftSetupCompleted;
  if (requestedTargets?.some((target) => !targetEnabled(target))) {
    throw new HttpError(409, "One or more requested calendar targets are disabled");
  }
  if (requestedTargets?.some((target) => !targetSetupCompleted(target))) {
    throw new HttpError(409, "One or more requested calendar targets have not completed setup and verification");
  }
  const configuredTargets = (["google", "microsoft"] as TargetProvider[]).filter((target) =>
    targetEnabled(target) && targetSetupCompleted(target),
  );
  const enabledTargets = requestedTargets?.length ? [...new Set(requestedTargets)] : configuredTargets;
  if (enabledTargets.length === 0) throw new HttpError(409, "No completed calendar target is enabled");
  if (enabledTargets.includes("google") && (!config.googleServiceAccountJson || !config.googleAdminEmail)) {
    throw new HttpError(409, "Google Workspace is not configured");
  }
  if (enabledTargets.includes("microsoft") && (!config.microsoftTenantId || !config.microsoftClientId || !config.microsoftClientSecret)) {
    throw new HttpError(409, "Microsoft 365 is not configured");
  }
  if (enabledTargets.includes("microsoft") && !config.microsoftConsentGrantedAt) {
    throw new HttpError(409, "Microsoft 365 admin consent has not been verified");
  }
  const schoolboxToken = config.schoolboxToken;

  await recoverStaleRuns();
  const run = await createRun(trigger);
  const heartbeat = setInterval(() => {
    void touchRunHeartbeat(run.id).catch(() => undefined);
  }, 30_000);
  const discoveryTimeoutMs = positiveTimeout(
    runtimeOptions.discoveryTimeoutMs,
    config.discoveryTimeoutSeconds * 1_000,
  );
  const userSyncTimeoutMs = positiveTimeout(
    runtimeOptions.userSyncTimeoutMs,
    config.userSyncTimeoutSeconds * 1_000,
  );
  const runTimeoutMs = positiveTimeout(
    runtimeOptions.runTimeoutMs,
    config.runTimeoutMinutes * 60_000,
  );

  try {
    await withHardDeadline(
      undefined,
      runTimeoutMs,
      `Organization synchronization timed out after ${durationLabel(runTimeoutMs)}.`,
      async (runSignal) => {
        await addAudit(actor, "sync.started", `Run ${run.id} started by ${trigger}`);
        const schoolbox = clientOverrides.schoolbox ?? new SchoolboxClient({
          baseUrl: config.schoolboxBaseUrl,
          jwt: schoolboxToken,
          pastDays: config.pastDays,
          futureDays: config.futureDays,
        });
        await checkpointRun(run, "discovery", `Waiting for Schoolbox and ${enabledTargets.length} calendar target director${enabledTargets.length === 1 ? "y" : "ies"}.`);
        let schoolboxProgress = "not started";
        const schoolboxUsers = await withHardDeadline(runSignal, discoveryTimeoutMs,
          `Schoolbox user discovery timed out after ${durationLabel(discoveryTimeoutMs)}.`,
          (signal) => schoolbox.getAllUsers({ signal, onPage: async (progress) => {
            schoolboxProgress = progress.totalItems === null
              ? `page ${progress.pageNumber}, ${progress.accumulatedItems} loaded`
              : `page ${progress.pageNumber}, ${progress.accumulatedItems} of ${progress.totalItems} loaded`;
            await checkpointRun(run, "discovery", `Schoolbox: ${schoolboxProgress}.`);
          } }));
        schoolboxProgress = "complete";
        const schoolboxByEmail = indexActiveSchoolboxUsersByEmail(schoolboxUsers);
        await checkpointRun(run, "target_sync", `Schoolbox discovery completed; ${enabledTargets.length} calendar target branch(es) are running.`);
        const targetResults = await Promise.all(enabledTargets.map(async (target) => {
          const targetRun = await startRunTarget(run.id, target);
          try {
            if (target === "google") {
              const google = clientOverrides.google ?? new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson!));
              targetRun.phase = "discovery";
              const users = await withHardDeadline(runSignal, discoveryTimeoutMs,
                `Google Directory user discovery timed out after ${durationLabel(discoveryTimeoutMs)}.`,
                (signal) => google.listAllUsers(config.googleAdminEmail, { customer: config.googleCustomer || "my_customer", signal,
                  onPage: async (progress) => checkpointRunTarget(targetRun, "discovery", `Google Directory page ${progress.pageNumber}: ${progress.accumulatedItems} loaded.`) }));
              const active = users.filter(isGoogleActive);
              const matched: MatchedUser[] = [];
              const discoveredAt = new Date().toISOString();
              const discoveries = active.map((user) => {
                const match = directorySchoolboxMatch(googleMatchEmails(user), schoolboxByEmail);
                if (match) matched.push({ google: user, schoolbox: match.schoolbox, schoolboxEmail: match.matchedEmail });
                return { target, targetUserId: user.id, targetEmail: user.primaryEmail, schoolboxUserId: match?.schoolbox.id ?? null,
                  schoolboxEmail: match?.matchedEmail ?? null, displayName: googleDisplayName(user) || (match ? schoolboxDisplayName(match.schoolbox) : null),
                  role: match ? schoolboxRole(match.schoolbox) : null, status: match ? "pending" : "unmatched",
                  lastSyncAt: null, lastError: match ? null : "No active Schoolbox user has a matching primary or alternate email address.", eventCount: 0, updatedAt: discoveredAt };
              });
              targetRun.usersDiscovered = active.length; targetRun.usersMatched = matched.length;
              const selection = await discoverUserMappings(discoveries, config.syncNewGoogleUsersByDefault, target);
              const selected = matched.filter((match) => selection.get(match.google.id)); targetRun.usersSelected = selected.length;
              await checkpointRunTarget(targetRun, "user_sync", `${selected.length} Google account(s) queued.`);
              let processed = 0;
              await processInPool(selected, config.concurrency, async (match) => {
                const deadline = createDeadlineSignal(runSignal, userSyncTimeoutMs, `User calendar synchronization timed out after ${durationLabel(userSyncTimeoutMs)}.`);
                try { await syncUser(match, run, schoolbox, google, { pastDays: config.pastDays, futureDays: config.futureDays, timezone: config.timezone, syncPolicy: config.syncPolicy, signal: deadline.signal }); }
                finally { deadline.dispose(); }
                processed += 1;
                await checkpointRunTarget(targetRun, "user_sync", `${processed} of ${selected.length} Google account(s) processed.`);
              }, runSignal);
              // Derive target totals from persisted per-user diagnostics. Reading the
              // completed rows avoids races between concurrent Google user workers.
              const diagnostics = await listRunUserDiagnostics(run.id, "google");
              targetRun.usersSynced = diagnostics.filter((item) => item.status === "completed").length;
              targetRun.errors = diagnostics.filter((item) => item.status === "failed").length;
              targetRun.eventsCreated = diagnostics.reduce((sum, item) => sum + item.eventsCreated, 0);
              targetRun.eventsUpdated = diagnostics.reduce((sum, item) => sum + item.eventsUpdated, 0);
              targetRun.eventsDeleted = diagnostics.reduce((sum, item) => sum + item.eventsDeleted, 0);
              targetRun.eventsUnchanged = diagnostics.reduce((sum, item) => sum + item.eventsUnchanged, 0);
              await checkpointRunTarget(targetRun, "user_sync", `${diagnostics.length} of ${selected.length} Google account(s) processed.`);
            } else {
              const microsoft = clientOverrides.microsoft ?? new MicrosoftGraphClient({ tenantId: config.microsoftTenantId, clientId: config.microsoftClientId, clientSecret: config.microsoftClientSecret! });
              const users = await withHardDeadline(runSignal, discoveryTimeoutMs,
                `Microsoft Entra user discovery timed out after ${durationLabel(discoveryTimeoutMs)}.`,
                (signal) => microsoft.listAllUsers({ signal, onPage: async (progress) => checkpointRunTarget(targetRun, "discovery", `Microsoft Entra page ${progress.pageNumber}: ${progress.accumulatedItems} loaded.`) }));
              const active = users.filter(isMicrosoftActive); const matched: MicrosoftMatchedUser[] = []; const discoveredAt = new Date().toISOString();
              const discoveries = active.map((user) => {
                const email = microsoftEmail(user); const match = directorySchoolboxMatch(microsoftMatchEmails(user), schoolboxByEmail);
                if (match) matched.push({ microsoft: user, schoolbox: match.schoolbox, schoolboxEmail: match.matchedEmail });
                return { target, targetUserId: user.id, targetEmail: email, schoolboxUserId: match?.schoolbox.id ?? null,
                  schoolboxEmail: match?.matchedEmail ?? null, displayName: microsoftDisplayName(user) || (match ? schoolboxDisplayName(match.schoolbox) : null),
                  role: match ? schoolboxRole(match.schoolbox) : null, status: match ? "pending" : "unmatched",
                  lastSyncAt: null, lastError: match ? null : "No active Schoolbox user has a matching Microsoft 365 address.", eventCount: 0, updatedAt: discoveredAt };
              });
              targetRun.usersDiscovered = active.length; targetRun.usersMatched = matched.length;
              const selection = await discoverUserMappings(discoveries, config.syncNewMicrosoftUsersByDefault, target);
              const selected = matched.filter((match) => selection.get(match.microsoft.id)); targetRun.usersSelected = selected.length;
              await checkpointRunTarget(targetRun, "user_sync", `${selected.length} Microsoft 365 account(s) queued.`);
              await processInPool(selected, config.concurrency, async (match) => {
                const deadline = createDeadlineSignal(runSignal, userSyncTimeoutMs, `Microsoft 365 user synchronization timed out after ${durationLabel(userSyncTimeoutMs)}.`);
                try { await syncMicrosoftUser(match, run, targetRun, schoolbox, microsoft, { pastDays: config.pastDays, futureDays: config.futureDays, timezone: config.timezone, syncPolicy: config.microsoftSyncPolicy, signal: deadline.signal }); }
                finally { deadline.dispose(); }
                await checkpointRunTarget(targetRun, "user_sync", `${targetRun.usersSynced + targetRun.errors} of ${selected.length} Microsoft 365 account(s) processed.`);
              }, runSignal);
            }
            targetRun.status = targetRun.errors ? "completed_with_errors" : "completed"; targetRun.phase = "completed";
            targetRun.message = `${targetRun.usersSynced} account(s) synced; ${targetRun.usersMatched - targetRun.usersSelected} matched account(s) paused.`;
          } catch (error) {
            targetRun.status = "failed"; targetRun.phase = "failed"; targetRun.errors += 1; targetRun.message = syncErrorMessage(error, runSignal);
          } finally { targetRun.completedAt = new Date().toISOString(); await finishRunTarget(targetRun); }
          return targetRun;
        }));
        for (const targetRun of targetResults) {
          run.usersDiscovered += targetRun.usersDiscovered; run.usersMatched += targetRun.usersMatched;
          run.usersSynced += targetRun.target === "microsoft" ? targetRun.usersSynced : 0;
          run.eventsCreated += targetRun.target === "microsoft" ? targetRun.eventsCreated : 0;
          run.eventsUpdated += targetRun.target === "microsoft" ? targetRun.eventsUpdated : 0;
          run.eventsDeleted += targetRun.target === "microsoft" ? targetRun.eventsDeleted : 0;
          run.eventsUnchanged += targetRun.target === "microsoft" ? targetRun.eventsUnchanged : 0;
          run.errors += targetRun.target === "microsoft" ? targetRun.errors : targetRun.status === "failed" ? 1 : 0;
        }
        throwIfAborted(runSignal);
        await checkpointRun(run, "finalizing", "Finalizing run counters and audit status.");
        run.status = run.errors > 0 ? "completed_with_errors" : "completed";
        const paused = targetResults.reduce((sum, target) => sum + target.usersMatched - target.usersSelected, 0);
        run.message = run.errors > 0
          ? `${run.errors} target/user sync(s) require attention; ${paused} matched target account(s) were paused.`
          : `Organization sync completed across ${enabledTargets.length} target(s); ${run.usersSynced} target account(s) synced and ${paused} paused.`;
      },
    );
  } catch (error) {
    run.status = "failed";
    run.errors += 1;
    run.message = syncErrorMessage(error);
  } finally {
    clearInterval(heartbeat);
    run.completedAt = new Date().toISOString();
    await finishRun(run);
    await addAudit(actor, `sync.${run.status}`, `Run ${run.id}: ${run.message ?? run.status}`);
  }

  return run;
}

export async function runScheduledSyncIfDue(actor: string): Promise<{
  status: "disabled" | "not_due" | "started";
  nextDueAt?: string;
  run?: RunSummary;
}> {
  await recoverStaleRuns();
  const config = await getConfig(false);
  if (!config.enabled || !config.setupCompleted) return { status: "disabled" };

  const lastRun = (await listRuns(1))[0];
  if (lastRun?.status === "running") return { status: "not_due" };
  if (lastRun) {
    const nextDue = new Date(lastRun.startedAt).getTime() + config.syncIntervalMinutes * 60_000;
    if (Date.now() < nextDue) {
      return { status: "not_due", nextDueAt: new Date(nextDue).toISOString() };
    }
  }

  return { status: "started", run: await runFullSync("scheduled", actor) };
}
