import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import { normalizeTargetProvider, type TargetProvider } from "@/lib/storage";
import { runFullSync } from "@/lib/sync";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const actor = await requestActor(request, "operate");
    const text = await request.text();
    let body: { targets?: unknown } = {};
    if (text) {
      try { body = JSON.parse(text) as { targets?: unknown }; }
      catch { throw new HttpError(400, "The sync request body is not valid JSON"); }
    }
    let targets: TargetProvider[] | undefined;
    if (body.targets !== undefined) {
      if (!Array.isArray(body.targets) || body.targets.length === 0 || body.targets.length > 2) {
        throw new HttpError(400, "Choose one or both calendar targets");
      }
      targets = [...new Set(body.targets.map(normalizeTargetProvider))];
    }
    return Response.json(await runFullSync("manual", actor, {}, {}, targets));
  } catch (error) {
    return jsonError(error);
  }
}
