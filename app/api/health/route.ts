import { db } from "@/lib/db";
import { deploymentReady } from "@/lib/auth";

export const dynamic = "force-dynamic";

export async function GET() {
  try {
    db().prepare("SELECT 1 AS ok").first<{ ok: number }>();
    const ready = deploymentReady();
    return Response.json({ ok: ready }, { status: ready ? 200 : 503, headers: { "Cache-Control": "no-store" } });
  } catch {
    return Response.json({ ok: false }, { status: 503, headers: { "Cache-Control": "no-store" } });
  }
}
