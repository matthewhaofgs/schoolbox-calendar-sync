import { requireSession } from "@/lib/auth";
import { listDiscoveredEventTypes } from "@/lib/storage";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    requireSession(request, "configure");
    return Response.json(
      { eventTypes: await listDiscoveredEventTypes() },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
