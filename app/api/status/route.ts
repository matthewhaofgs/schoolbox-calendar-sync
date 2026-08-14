import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { listRunTargets, statusSnapshot } from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const snapshot = await statusSnapshot();
    const lastRunTargets = snapshot.lastRun ? await listRunTargets(snapshot.lastRun.id) : [];
    return Response.json({
      configured: snapshot.configured,
      lastRun: snapshot.lastRun ? { ...snapshot.lastRun, targets: lastRunTargets } : null,
      lastRunTargets,
      counts: snapshot.counts,
      targetCounts: snapshot.targetCounts,
      schedule: {
        enabled: snapshot.config.enabled,
        setupCompleted: snapshot.config.setupCompleted,
        schoolboxSetupCompleted: snapshot.config.schoolboxSetupCompleted,
        googleSetupCompleted: snapshot.config.googleSetupCompleted,
        microsoftSetupCompleted: snapshot.config.microsoftSetupCompleted,
        schoolboxConfigured: snapshot.config.schoolboxConfigured,
        googleConfigured: snapshot.config.googleConfigured,
        microsoftConfigured: snapshot.config.microsoftConfigured,
        syncIntervalMinutes: snapshot.config.syncIntervalMinutes,
        syncNewUsersByDefault: snapshot.config.syncNewUsersByDefault,
        googleEnabled: snapshot.config.googleEnabled,
        syncNewGoogleUsersByDefault: snapshot.config.syncNewGoogleUsersByDefault,
        microsoftEnabled: snapshot.config.microsoftEnabled,
        syncNewMicrosoftUsersByDefault: snapshot.config.syncNewMicrosoftUsersByDefault,
        microsoftConsentGrantedAt: snapshot.config.microsoftConsentGrantedAt,
        pastDays: snapshot.config.pastDays,
        futureDays: snapshot.config.futureDays,
        discoveryTimeoutSeconds: snapshot.config.discoveryTimeoutSeconds,
        userSyncTimeoutSeconds: snapshot.config.userSyncTimeoutSeconds,
        runTimeoutMinutes: snapshot.config.runTimeoutMinutes,
        timezone: snapshot.config.timezone,
        syncPolicy: snapshot.config.syncPolicy,
        microsoftSyncPolicy: snapshot.config.microsoftSyncPolicy,
      },
    });
  } catch (error) {
    return jsonError(error);
  }
}
