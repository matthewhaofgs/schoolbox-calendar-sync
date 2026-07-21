import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import { getConfig, listCalendarDestinationUsage } from "@/lib/storage";
import { retireCalendarDestination } from "@/lib/sync";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "configure");
    return Response.json({ destinations: await listCalendarDestinationUsage() });
  } catch (error) {
    return jsonError(error);
  }
}

export async function DELETE(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const body = await request.json() as { destinationId?: unknown };
    if (typeof body.destinationId !== "string" || !body.destinationId.trim()) {
      throw new HttpError(400, "Choose a calendar destination to retire");
    }
    const result = await retireCalendarDestination(body.destinationId, actor);
    return Response.json({
      ...result,
      config: await getConfig(false),
      destinations: await listCalendarDestinationUsage(),
    });
  } catch (error) {
    return jsonError(error);
  }
}
