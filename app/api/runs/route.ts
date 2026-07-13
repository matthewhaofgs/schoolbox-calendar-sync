import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { listRuns } from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const limit = Number(new URL(request.url).searchParams.get("limit") ?? 30);
    return Response.json({ runs: await listRuns(limit) });
  } catch (error) {
    return jsonError(error);
  }
}
