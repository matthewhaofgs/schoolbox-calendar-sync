import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import {
  getEventMappings,
  getUserMapping,
  listUserCalendarTargets,
  listUserRunDiagnostics,
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

    const [events, calendars, runs] = await Promise.all([
      getEventMappings(userId),
      listUserCalendarTargets(userId),
      listUserRunDiagnostics(userId, 20),
    ]);
    return privateJson({ user, events, calendars, runs });
  } catch (error) {
    return jsonError(error);
  }
}
