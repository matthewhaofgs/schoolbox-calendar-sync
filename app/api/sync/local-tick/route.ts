import { requestScheduler } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { runScheduledSyncIfDue } from "@/lib/sync";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const actor = requestScheduler(request);
    return Response.json(await runScheduledSyncIfDue(actor));
  } catch (error) {
    return jsonError(error);
  }
}
