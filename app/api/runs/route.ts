import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { listRuns, listRunTargets } from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const limit = Number(new URL(request.url).searchParams.get("limit") ?? 30);
    const runs = await listRuns(limit);
    return Response.json({
      runs: await Promise.all(runs.map(async (run) => ({ ...run, targets: await listRunTargets(run.id) }))),
    });
  } catch (error) {
    return jsonError(error);
  }
}
