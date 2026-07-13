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
  addAudit,
  createRun,
  deleteEventMapping,
  discoverUserMappings,
  finishRun,
  getConfig,
  getEventMappings,
  listRuns,
  recoverStaleRuns,
  touchRunHeartbeat,
  touchEventMapping,
  upsertEventMapping,
  upsertUserMapping,
  type RunSummary,
} from "./storage";
import { HttpError } from "./security";

type MatchedUser = { google: GoogleDirectoryUser; schoolbox: SchoolboxUser };
type SchoolboxSyncClient = Pick<SchoolboxClient, "getAllUsers" | "getCalendarEvents">;
type GoogleSyncClient = Pick<GoogleWorkspaceClient, "listAllUsers" | "insertEvent" | "updateEvent" | "deleteEvent">;

/** Optional client overrides used by deterministic integration tests. */
export type SyncClientOverrides = {
  schoolbox?: SchoolboxSyncClient;
  google?: GoogleSyncClient;
};

function normalizedEmail(value: string | null | undefined): string {
  return (value ?? "").trim().toLowerCase();
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

async function eventBody(
  event: NormalizedSchoolboxCalendarEvent,
  googleUserId: string,
  timezone: string,
  sourceKey: string,
): Promise<GoogleCalendarEventInput & { id: string }> {
  const id = await createDeterministicEventId(`${googleUserId}:${sourceKey}`);
  const sourceLink = event.sourceUrl;
  const descriptionParts = [event.description, sourceLink ? `Schoolbox: ${sourceLink}` : undefined].filter(Boolean);
  const body: GoogleCalendarEventInput & { id: string } = {
    id,
    summary: event.title || "Schoolbox event",
    description: descriptionParts.join("\n\n") || undefined,
    location: event.location || undefined,
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
  if (sourceLink) body.source = { title: "Open in Schoolbox", url: sourceLink };
  return body;
}

async function processInPool<T>(items: T[], concurrency: number, worker: (item: T) => Promise<void>): Promise<void> {
  let index = 0;
  const runners = Array.from({ length: Math.min(Math.max(1, concurrency), Math.max(1, items.length)) }, async () => {
    while (index < items.length) {
      const item = items[index];
      index += 1;
      await worker(item);
    }
  });
  await Promise.all(runners);
}

async function syncUser(
  match: MatchedUser,
  run: RunSummary,
  schoolbox: SchoolboxSyncClient,
  google: GoogleSyncClient,
  options: { pastDays: number; futureDays: number; timezone: string },
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
    schoolboxEmail: match.schoolbox.email ?? null,
    displayName: googleDisplayName(match.google) || schoolboxDisplayName(match.schoolbox),
    role: schoolboxRole(match.schoolbox),
    updatedAt: new Date().toISOString(),
  };

  try {
    const [events, storedMappings] = await Promise.all([
      schoolbox.getCalendarEvents(match.schoolbox.id, {
        pastDays: options.pastDays,
        futureDays: options.futureDays,
        now,
      }),
      getEventMappings(googleUserId),
    ]);
    const existing = new Map(storedMappings.map((mapping) => [mapping.sourceKey, mapping]));
    const seen = new Set<string>();

    for (const event of events) {
      const sourceKey = `${event.sourceKey}:occurrence:${event.start}`;
      if (seen.has(sourceKey)) continue;
      seen.add(sourceKey);
      const body = await eventBody(event, googleUserId, options.timezone, sourceKey);
      const hash = await createContentHash(body);
      const mapping = existing.get(sourceKey);

      if (mapping?.sourceHash === hash) {
        await touchEventMapping(googleUserId, sourceKey, run.id);
        run.eventsUnchanged += 1;
        continue;
      }

      let createdAt = mapping?.createdAt ?? new Date().toISOString();
      if (mapping) {
        await google.updateEvent(googleEmail, mapping.googleEventId, body, {
          calendarId: "primary",
          quotaUser: googleUserId,
          sendUpdates: "none",
        });
        run.eventsUpdated += 1;
      } else {
        try {
          await google.insertEvent(googleEmail, body, {
            calendarId: "primary",
            quotaUser: googleUserId,
            sendUpdates: "none",
          });
          run.eventsCreated += 1;
        } catch (error) {
          if (!(error instanceof GoogleApiError) || error.status !== 409) throw error;
          await google.updateEvent(googleEmail, body.id, body, {
            calendarId: "primary",
            quotaUser: googleUserId,
            sendUpdates: "none",
          });
          run.eventsUpdated += 1;
          createdAt = new Date().toISOString();
        }
      }

      await upsertEventMapping({
        googleUserId,
        sourceKey,
        googleEventId: mapping?.googleEventId ?? body.id,
        sourceHash: hash,
        sourceStart: event.start,
        sourceEnd: event.end,
        lastSeenRunId: run.id,
        createdAt,
        updatedAt: new Date().toISOString(),
      });
    }

    for (const mapping of storedMappings) {
      if (seen.has(mapping.sourceKey)) continue;
      const sourceStart = new Date(mapping.sourceStart);
      const sourceEnd = new Date(mapping.sourceEnd);
      const wasInsideFetchedWindow = sourceStart < windowEnd && sourceEnd > windowStart;
      if (!wasInsideFetchedWindow) continue;
      try {
        await google.deleteEvent(googleEmail, mapping.googleEventId, {
          calendarId: "primary",
          quotaUser: googleUserId,
          sendUpdates: "none",
        });
      } catch (error) {
        if (!(error instanceof GoogleApiError) || (error.status !== 404 && error.status !== 410)) throw error;
      }
      await deleteEventMapping(googleUserId, mapping.sourceKey);
      run.eventsDeleted += 1;
    }

    await upsertUserMapping({
      ...baseMapping,
      status: "synced",
      lastSyncAt: new Date().toISOString(),
      lastError: null,
      eventCount: events.length,
    });
    run.usersSynced += 1;
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown per-user sync error";
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

export async function runFullSync(
  trigger: string,
  actor: string,
  clientOverrides: SyncClientOverrides = {},
): Promise<RunSummary> {
  const config = await getConfig(true);
  if (!config.schoolboxBaseUrl || !config.schoolboxToken) {
    throw new HttpError(409, "Schoolbox is not configured");
  }
  if (!config.googleServiceAccountJson || !config.googleAdminEmail) {
    throw new HttpError(409, "Google Workspace is not configured");
  }

  await recoverStaleRuns();
  const run = await createRun(trigger);
  const heartbeat = setInterval(() => {
    void touchRunHeartbeat(run.id).catch(() => undefined);
  }, 30_000);

  try {
    await addAudit(actor, "sync.started", `Run ${run.id} started by ${trigger}`);
    const schoolbox = clientOverrides.schoolbox ?? new SchoolboxClient({
      baseUrl: config.schoolboxBaseUrl,
      jwt: config.schoolboxToken,
      pastDays: config.pastDays,
      futureDays: config.futureDays,
    });
    const google = clientOverrides.google
      ?? new GoogleWorkspaceClient(parseServiceAccountJson(config.googleServiceAccountJson));
    const [schoolboxUsers, googleUsers] = await Promise.all([
      schoolbox.getAllUsers(),
      google.listAllUsers(config.googleAdminEmail, { customer: config.googleCustomer || "my_customer" }),
    ]);

    const activeSchoolbox = schoolboxUsers.filter(isSchoolboxActive);
    const schoolboxByEmail = new Map(
      activeSchoolbox.filter((user) => normalizedEmail(user.email)).map((user) => [normalizedEmail(user.email), user]),
    );
    const activeGoogle = googleUsers.filter(isGoogleActive);
    const matched: MatchedUser[] = [];
    const discoveredAt = new Date().toISOString();
    const discoveries = activeGoogle.map((googleUser) => {
      const schoolboxUser = schoolboxByEmail.get(normalizedEmail(googleUser.primaryEmail));
      if (schoolboxUser) matched.push({ google: googleUser, schoolbox: schoolboxUser });
      return {
        googleUserId: googleUser.id,
        googleEmail: googleUser.primaryEmail,
        schoolboxUserId: schoolboxUser?.id ?? null,
        schoolboxEmail: schoolboxUser?.email ?? null,
        displayName: googleDisplayName(googleUser) || (schoolboxUser ? schoolboxDisplayName(schoolboxUser) : null),
        role: schoolboxUser ? schoolboxRole(schoolboxUser) : null,
        status: schoolboxUser ? "pending" : "unmatched",
        lastSyncAt: null,
        lastError: schoolboxUser ? null : "No active Schoolbox user has this primary email address.",
        eventCount: 0,
        updatedAt: discoveredAt,
      };
    });
    run.usersDiscovered = activeGoogle.length;
    const selection = await discoverUserMappings(discoveries, config.syncNewUsersByDefault);
    run.usersMatched = matched.length;
    const selected = matched.filter((match) => selection.get(match.google.id) === true);
    await processInPool(selected, config.concurrency, (match) =>
      syncUser(match, run, schoolbox, google, {
        pastDays: config.pastDays,
        futureDays: config.futureDays,
        timezone: config.timezone,
      }),
    );
    run.status = run.errors > 0 ? "completed_with_errors" : "completed";
    const paused = matched.length - selected.length;
    run.message = run.errors > 0
      ? `${run.errors} user syncs require attention; ${paused} matched user(s) were paused.`
      : `Organization sync completed; ${run.usersSynced} user(s) synced and ${paused} matched user(s) paused.`;
  } catch (error) {
    run.status = "failed";
    run.errors += 1;
    run.message = error instanceof Error ? error.message : "The organization sync failed.";
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
