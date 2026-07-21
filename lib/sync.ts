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
  eventIncludedByPolicy,
  resolveGoogleEventRule,
  withoutManagedCalendarDestination,
  type SyncPolicy,
} from "./policy";
import {
  addAudit,
  checkpointRun,
  createRun,
  deleteCalendarTargetRecords,
  deleteEventMapping,
  discoverUserMappings,
  finishRun,
  getConfig,
  getEventMappings,
  getUserCalendarTarget,
  getUserMapping,
  listCalendarTargetsForDestination,
  listUserCalendarTargets,
  listRuns,
  recordManagedEventCleanup,
  recoverStaleRuns,
  recordDiscoveredEventTypes,
  saveConfig,
  setUsersSyncEnabled,
  touchRunHeartbeat,
  touchEventMapping,
  upsertEventMapping,
  upsertUserCalendarTarget,
  upsertUserMapping,
  type RunSummary,
  type UserCalendarTarget,
} from "./storage";
import { HttpError } from "./security";

type MatchedUser = { google: GoogleDirectoryUser; schoolbox: SchoolboxUser; schoolboxEmail: string };
type SchoolboxSyncClient = Pick<SchoolboxClient, "getAllUsers" | "getCalendarEvents">;
type GoogleSyncClient = Pick<GoogleWorkspaceClient, "listAllUsers" | "createCalendar" | "updateCalendar" | "insertEvent" | "updateEvent" | "deleteEvent">;
type GoogleCleanupClient = Pick<GoogleWorkspaceClient, "deleteEvent"> & Partial<Pick<GoogleWorkspaceClient, "deleteCalendar">>;
type GoogleCalendarRetirementClient = Pick<GoogleWorkspaceClient, "deleteCalendar">;

/** Optional client overrides used by deterministic integration tests. */
export type SyncClientOverrides = {
  schoolbox?: SchoolboxSyncClient;
  google?: GoogleSyncClient;
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
  const runners = Array.from({ length: Math.min(Math.max(1, concurrency), Math.max(1, items.length)) }, async () => {
    while (index < items.length) {
      throwIfAborted(signal);
      const item = items[index];
      index += 1;
      await worker(item);
    }
  });
  await Promise.all(runners);
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

  try {
    const [events, storedMappings, storedCalendarTargets] = await Promise.all([
      schoolbox.getCalendarEvents(match.schoolbox.id, {
        pastDays: options.pastDays,
        futureDays: options.futureDays,
        now,
        signal: options.signal,
      }),
      getEventMappings(googleUserId),
      listUserCalendarTargets(googleUserId),
    ]);
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
    const calendarTargets = new Map<string, Promise<string>>();
    await recordDiscoveredEventTypes(events);

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

    for (const event of events) {
      throwIfAborted(options.signal);
      const sourceKey = `${event.sourceKey}:occurrence:${event.start}`;
      if (seen.has(sourceKey) || excluded.has(sourceKey)) continue;
      if (!eventIncludedByPolicy({
        category: event.category ?? "other",
        type: event.type,
        allDay: event.allDay,
        completed: Boolean(event.completed),
      }, options.syncPolicy)) {
        excluded.add(sourceKey);
        excludedSourceRoots.add(event.sourceKey);
        continue;
      }
      seen.add(sourceKey);
      const googleRule = resolveGoogleEventRule({ category: event.category ?? "other", type: event.type }, options.syncPolicy);
      const calendarId = await targetFor(googleRule.destinationId);
      const body = await eventBody(event, googleUserId, options.timezone, sourceKey, options.syncPolicy);
      const hash = await createContentHash(body);
      const mapping = existing.get(sourceKey);
      const calendarChanged = Boolean(mapping && mapping.calendarId !== calendarId);

      if (mapping?.sourceHash === hash && !calendarChanged) {
        await touchEventMapping(googleUserId, sourceKey, run.id);
        run.eventsUnchanged += 1;
        continue;
      }

      let createdAt = mapping?.createdAt ?? new Date().toISOString();
      if (mapping && !calendarChanged) {
        await google.updateEvent(googleEmail, mapping.googleEventId, body, {
          calendarId,
          quotaUser: googleUserId,
          sendUpdates: "none",
          signal: options.signal,
        });
        run.eventsUpdated += 1;
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
        } else {
          if (insertConflict) run.eventsUpdated += 1;
          else run.eventsCreated += 1;
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
      });
    }

    for (const mapping of storedMappings) {
      throwIfAborted(options.signal);
      if (seen.has(mapping.sourceKey)) continue;
      const occurrenceMarker = mapping.sourceKey.lastIndexOf(":occurrence:");
      const mappingRoot = occurrenceMarker >= 0 ? mapping.sourceKey.slice(0, occurrenceMarker) : mapping.sourceKey;
      const excludedByPolicy = excluded.has(mapping.sourceKey) || excludedSourceRoots.has(mappingRoot);
      if (excludedByPolicy ? !options.syncPolicy.deleteExcludedEvents : !options.syncPolicy.deleteMissingEvents) continue;
      const sourceStart = new Date(mapping.sourceStart);
      const sourceEnd = new Date(mapping.sourceEnd);
      const wasInsideFetchedWindow = sourceStart < windowEnd && sourceEnd > windowStart;
      if (!wasInsideFetchedWindow) continue;
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
      await deleteEventMapping(googleUserId, mapping.sourceKey);
      run.eventsDeleted += 1;
    }

    throwIfAborted(options.signal);
    const managedEventCount = (await getEventMappings(googleUserId)).length;
    await upsertUserMapping({
      ...baseMapping,
      status: "synced",
      lastSyncAt: new Date().toISOString(),
      lastError: null,
      eventCount: managedEventCount,
    });
    run.usersSynced += 1;
  } catch (error) {
    const message = syncErrorMessage(error, options.signal);
    await upsertUserMapping({
      ...baseMapping,
      status: "error",
      lastSyncAt: new Date().toISOString(),
      lastError: message.slice(0, 2000),
      eventCount: 0,
    });
    run.errors += 1;
  }
}

/**
 * Pauses one user and removes Google events tracked in Relay's event_mappings
 * table. When explicitly requested, it then permanently deletes only the
 * secondary calendars recorded in user_calendar_targets. Deleting a calendar
 * also deletes any manually added content it contains, so this path never
 * accepts a primary or untracked calendar ID.
 */
export async function cleanupUserManagedEvents(
  googleUserId: string,
  actor: string,
  clientOverride?: GoogleCleanupClient,
  options: { deleteCalendars?: boolean } = {},
): Promise<ManagedEventCleanupResult> {
  const userId = googleUserId.trim();
  if (!userId) throw new HttpError(400, "Choose a user to clean up");

  const mapping = await getUserMapping(userId);
  if (!mapping) throw new HttpError(404, "This user is no longer available");

  const [storedMappings, storedCalendarTargets] = await Promise.all([
    getEventMappings(userId),
    listUserCalendarTargets(userId),
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
  await setUsersSyncEnabled([userId], false, actor);
  if (!hasDestructiveWork) {
    await recordManagedEventCleanup({
      googleUserId: userId,
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

  let google: GoogleCleanupClient;
  try {
    if (clientOverride) {
      google = clientOverride;
    } else {
      const config = await getConfig(true);
      if (!config.googleServiceAccountJson) throw new HttpError(409, "Google Workspace is not configured");
      google = new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson));
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Google Workspace cleanup could not start";
    await recordManagedEventCleanup({
      googleUserId: userId,
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
  for (const eventMapping of storedMappings) {
    try {
      await google.deleteEvent(mapping.googleEmail, eventMapping.googleEventId, {
        calendarId: eventMapping.calendarId,
        quotaUser: userId,
        sendUpdates: "none",
      });
      deleted += 1;
    } catch (error) {
      if (error instanceof GoogleApiError && (error.status === 404 || error.status === 410)) {
        alreadyMissing += 1;
      } else {
        cleanupError = error instanceof Error ? error.message : "Google Calendar cleanup failed";
        break;
      }
    }
    await deleteEventMapping(userId, eventMapping.sourceKey);
  }

  let remaining = (await getEventMappings(userId)).length;
  let calendarsDeleted = 0;
  let calendarsAlreadyMissing = 0;
  if (deleteCalendars && remaining === 0 && storedCalendarTargets.length > 0) {
    if (!google.deleteCalendar) {
      cleanupError = "Google Calendar cleanup does not support deleting secondary calendars";
    } else {
      for (const target of storedCalendarTargets) {
        let removeTargetRecords = false;
        try {
          await google.deleteCalendar(mapping.googleEmail, target.googleCalendarId, { quotaUser: userId });
          calendarsDeleted += 1;
          removeTargetRecords = true;
        } catch (error) {
          if (error instanceof GoogleApiError && (error.status === 404 || error.status === 410)) {
            calendarsAlreadyMissing += 1;
            removeTargetRecords = true;
          } else {
            cleanupError = error instanceof Error ? error.message : "Google Calendar cleanup failed";
          }
        }
        if (removeTargetRecords) await deleteCalendarTargetRecords(userId, target.destinationId, target.googleCalendarId);
      }
      remaining = (await getEventMappings(userId)).length;
    }
  }

  const calendarsRemaining = (await listUserCalendarTargets(userId)).length;
  await recordManagedEventCleanup({
    googleUserId: userId,
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
 * tracked user. The destination is removed from policy before Google deletion
 * starts so a subsequent sync cannot recreate it. Failed target records are
 * retained and shown as retired cleanup work that an administrator can retry.
 */
export async function retireCalendarDestination(
  destinationId: string,
  actor: string,
  clientOverride?: GoogleCalendarRetirementClient,
): Promise<CalendarDestinationRetirementResult> {
  const id = destinationId.trim();
  if (!/^[a-z0-9][a-z0-9_-]{0,39}$/.test(id)) throw new HttpError(400, "Choose a valid calendar destination");
  if (id === "primary") throw new HttpError(400, "The primary calendar cannot be deleted");

  const initialConfig = await getConfig(false);
  const initialTargets = await listCalendarTargetsForDestination(id);
  if (!initialConfig.syncPolicy.secondaryCalendars.some((calendar) => calendar.id === id) && initialTargets.length === 0) {
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
    const targets = await listCalendarTargetsForDestination(id);
    run.usersDiscovered = targets.length;
    run.usersMatched = targets.length;
    await checkpointRun(run, "calendar_retirement", `${targets.length} tracked user calendar(s) queued for retirement.`);

    let google: GoogleCalendarRetirementClient | null = clientOverride ?? null;
    if (targets.length > 0 && !google) {
      if (!config.googleServiceAccountJson) throw new HttpError(409, "Google Workspace is not configured");
      google = new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson));
    }

    if (config.syncPolicy.secondaryCalendars.some((calendar) => calendar.id === id)) {
      await saveConfig({
        syncPolicy: withoutManagedCalendarDestination(config.syncPolicy, id),
      }, actor);
    }

    const errors: string[] = [];
    await processInPool(targets, config.concurrency, async (target) => {
      let removeTargetRecords = false;
      try {
        await google!.deleteCalendar(target.googleEmail, target.googleCalendarId, {
          quotaUser: target.googleUserId,
        });
        result.calendarsDeleted += 1;
        removeTargetRecords = true;
      } catch (error) {
        if (error instanceof GoogleApiError && (error.status === 404 || error.status === 410)) {
          result.calendarsAlreadyMissing += 1;
          removeTargetRecords = true;
        } else {
          result.calendarsFailed += 1;
          const message = error instanceof Error ? error.message : "Google Calendar deletion failed";
          if (!errors.includes(message)) errors.push(message);
        }
      }
      if (removeTargetRecords) {
        const removed = await deleteCalendarTargetRecords(
          target.googleUserId,
          id,
          target.googleCalendarId,
        );
        result.eventMappingsRemoved += removed;
        run.eventsDeleted += removed;
        run.usersSynced += 1;
      }
      await checkpointRun(
        run,
        "calendar_retirement",
        `${run.usersSynced + result.calendarsFailed} of ${targets.length} tracked user calendar(s) processed.`,
      );
    });

    result.calendarsRemaining = (await listCalendarTargetsForDestination(id)).length;
    result.error = errors.length > 0 ? errors.join("; ").slice(0, 2_000) : null;
    run.errors = result.calendarsFailed;
    run.status = result.calendarsFailed > 0 ? "completed_with_errors" : "completed";
    run.message = result.calendarsFailed > 0
      ? `Calendar destination ${id} was retired, but ${result.calendarsRemaining} user calendar(s) still require deletion.`
      : `Calendar destination ${id} was retired and ${result.calendarsDeleted + result.calendarsAlreadyMissing} tracked user calendar(s) were removed.`;
  } catch (error) {
    unexpectedError = error;
    run.status = "failed";
    run.errors += 1;
    run.message = error instanceof Error ? error.message : "Calendar destination retirement failed.";
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
): Promise<RunSummary> {
  const config = await getConfig(true);
  if (!config.schoolboxBaseUrl || !config.schoolboxToken) {
    throw new HttpError(409, "Schoolbox is not configured");
  }
  if (!config.googleServiceAccountJson || !config.googleAdminEmail) {
    throw new HttpError(409, "Google Workspace is not configured");
  }
  const schoolboxToken = config.schoolboxToken;
  const googleServiceAccountJson = config.googleServiceAccountJson;

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
        const google = clientOverrides.google
          ?? new GoogleWorkspaceClient(parseServiceAccountJson(googleServiceAccountJson));
        await checkpointRun(run, "discovery", "Waiting for Schoolbox and Google Directory user lists.");
        const discoveryController = new AbortController();
        const forwardRunAbort = () => discoveryController.abort(runSignal.reason);
        if (runSignal.aborted) forwardRunAbort();
        else runSignal.addEventListener("abort", forwardRunAbort, { once: true });
        let schoolboxProgress = "not started";
        let googleProgress = "not started";
        const reportDiscoveryProgress = () => checkpointRun(
          run,
          "discovery",
          `Schoolbox: ${schoolboxProgress}. Google Directory: ${googleProgress}.`,
        );
        const [schoolboxUsers, googleUsers] = await (async () => {
          try {
            const schoolboxPromise = withHardDeadline(
              discoveryController.signal,
              discoveryTimeoutMs,
              `Schoolbox user discovery timed out after ${durationLabel(discoveryTimeoutMs)}.`,
              (signal) => schoolbox.getAllUsers({
                signal,
                onPage: async (progress) => {
                  schoolboxProgress = progress.totalItems === null
                    ? `page ${progress.pageNumber}, ${progress.accumulatedItems} loaded`
                    : `page ${progress.pageNumber}, ${progress.accumulatedItems} of ${progress.totalItems} loaded`;
                  await reportDiscoveryProgress();
                },
              }),
            ).then(async (users) => {
              schoolboxProgress = "complete";
              await reportDiscoveryProgress();
              return users;
            });
            const googlePromise = withHardDeadline(
              discoveryController.signal,
              discoveryTimeoutMs,
              `Google Directory user discovery timed out after ${durationLabel(discoveryTimeoutMs)}.`,
              (signal) => google.listAllUsers(config.googleAdminEmail, {
                customer: config.googleCustomer || "my_customer",
                signal,
                onPage: async (progress) => {
                  googleProgress = `page ${progress.pageNumber}, ${progress.accumulatedItems} loaded`;
                  await reportDiscoveryProgress();
                },
              }),
            ).then(async (users) => {
              googleProgress = "complete";
              await reportDiscoveryProgress();
              return users;
            });
            return await Promise.all([schoolboxPromise, googlePromise]);
          } catch (error) {
            discoveryController.abort(error);
            throw error;
          } finally {
            runSignal.removeEventListener("abort", forwardRunAbort);
          }
        })();

        await checkpointRun(run, "matching", "Matching active Google and Schoolbox identities.");
        const schoolboxByEmail = indexActiveSchoolboxUsersByEmail(schoolboxUsers);
        const activeGoogle = googleUsers.filter(isGoogleActive);
        const matched: MatchedUser[] = [];
        const discoveredAt = new Date().toISOString();
        const discoveries = activeGoogle.map((googleUser) => {
          const googleEmail = normalizedEmail(googleUser.primaryEmail);
          const schoolboxUser = schoolboxByEmail.get(googleEmail);
          if (schoolboxUser) matched.push({ google: googleUser, schoolbox: schoolboxUser, schoolboxEmail: googleEmail });
          return {
            googleUserId: googleUser.id,
            googleEmail: googleUser.primaryEmail,
            schoolboxUserId: schoolboxUser?.id ?? null,
            schoolboxEmail: schoolboxUser ? googleEmail : null,
            displayName: googleDisplayName(googleUser) || (schoolboxUser ? schoolboxDisplayName(schoolboxUser) : null),
            role: schoolboxUser ? schoolboxRole(schoolboxUser) : null,
            status: schoolboxUser ? "pending" : "unmatched",
            lastSyncAt: null,
            lastError: schoolboxUser ? null : "No active Schoolbox user has this primary or alternate email address.",
            eventCount: 0,
            updatedAt: discoveredAt,
          };
        });
        run.usersDiscovered = activeGoogle.length;
        const selection = await discoverUserMappings(discoveries, config.syncNewUsersByDefault);
        run.usersMatched = matched.length;
        const selected = matched.filter((match) => selection.get(match.google.id) === true);
        await checkpointRun(
          run,
          "user_sync",
          `${selected.length} enabled user calendar(s) queued; ${matched.length - selected.length} matched user(s) paused.`,
        );
        await processInPool(selected, config.concurrency, async (match) => {
          const deadline = createDeadlineSignal(
            runSignal,
            userSyncTimeoutMs,
            `User calendar synchronization timed out after ${durationLabel(userSyncTimeoutMs)}.`,
          );
          try {
            await syncUser(match, run, schoolbox, google, {
              pastDays: config.pastDays,
              futureDays: config.futureDays,
              timezone: config.timezone,
              syncPolicy: config.syncPolicy,
              signal: deadline.signal,
            });
          } finally {
            deadline.dispose();
          }
          throwIfAborted(runSignal);
          await checkpointRun(
            run,
            "user_sync",
            `${run.usersSynced + run.errors} of ${selected.length} enabled user calendar(s) processed.`,
          );
        }, runSignal);
        throwIfAborted(runSignal);
        await checkpointRun(run, "finalizing", "Finalizing run counters and audit status.");
        run.status = run.errors > 0 ? "completed_with_errors" : "completed";
        const paused = matched.length - selected.length;
        run.message = run.errors > 0
          ? `${run.errors} user syncs require attention; ${paused} matched user(s) were paused.`
          : `Organization sync completed; ${run.usersSynced} user(s) synced and ${paused} matched user(s) paused.`;
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
