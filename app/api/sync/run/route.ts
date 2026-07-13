import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { runFullSync } from "@/lib/sync";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const actor = await requestActor(request, "operate");
    return Response.json(await runFullSync("manual", actor));
  } catch (error) {
    return jsonError(error);
  }
}
