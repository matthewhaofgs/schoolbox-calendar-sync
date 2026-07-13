import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { statusSnapshot } from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const snapshot = await statusSnapshot();
    return Response.json({
      configured: snapshot.configured,
      lastRun: snapshot.lastRun,
      counts: snapshot.counts,
      schedule: {
        enabled: snapshot.config.enabled,
        setupCompleted: snapshot.config.setupCompleted,
        syncIntervalMinutes: snapshot.config.syncIntervalMinutes,
        syncNewUsersByDefault: snapshot.config.syncNewUsersByDefault,
        pastDays: snapshot.config.pastDays,
        futureDays: snapshot.config.futureDays,
        timezone: snapshot.config.timezone,
      },
    });
  } catch (error) {
    return jsonError(error);
  }
}
