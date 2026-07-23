import { requestActor } from "@/lib/auth";
import { EVENT_CATEGORIES, type EventCategory } from "@/lib/policy";
import { HttpError, jsonError } from "@/lib/security";
import {
  getConfig,
  getEventMappings,
  getUserEventExclusions,
  getUserMapping,
  listDiscoveredEventTypes,
  listUserCalendarTargets,
  listUserRunDiagnostics,
  saveUserEventExclusions,
} from "@/lib/storage";

export const dynamic = "force-dynamic";

const privateJson = (body: unknown) => Response.json(body, {
  headers: { "Cache-Control": "private, no-store" },
});

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const userId = new URL(request.url).searchParams.get("userId")?.trim();
    if (!userId || userId.length > 200) throw new HttpError(400, "Choose a valid user");
    const user = await getUserMapping(userId);
    if (!user) throw new HttpError(404, "User not found");

    const [events, calendars, runs, exclusions, eventTypes, config] = await Promise.all([
      getEventMappings(userId),
      listUserCalendarTargets(userId),
      listUserRunDiagnostics(userId, 20),
      getUserEventExclusions(userId),
      listDiscoveredEventTypes(),
      getConfig(false),
    ]);
    return privateJson({ user, events, calendars, runs, exclusions, eventTypes, globalPolicy: config.syncPolicy });
  } catch (error) {
    return jsonError(error);
  }
}

export async function PATCH(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const body = await request.json() as {
      userId?: unknown;
      excludedCategories?: unknown;
      excludedEventTypes?: unknown;
    };
    if (typeof body.userId !== "string" || !body.userId.trim()) throw new HttpError(400, "Choose a user");
    if (!Array.isArray(body.excludedCategories) ||
      !body.excludedCategories.every((value) => typeof value === "string" && EVENT_CATEGORIES.includes(value as EventCategory))) {
      throw new HttpError(400, "Excluded categories must be valid event categories");
    }
    if (!Array.isArray(body.excludedEventTypes) ||
      !body.excludedEventTypes.every((value) => typeof value === "string" && value.trim().length > 0 && value.length <= 120)) {
      throw new HttpError(400, "Excluded event types must be non-empty labels");
    }
    if (body.excludedEventTypes.length > 200) throw new HttpError(400, "Exclude no more than 200 exact event types");
    const exclusions = await saveUserEventExclusions(body.userId, {
      categories: body.excludedCategories,
      eventTypes: body.excludedEventTypes,
    }, actor);
    return privateJson({ exclusions });
  } catch (error) {
    return jsonError(error);
  }
}
