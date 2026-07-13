import { getOAuthSettings, requireSession, saveOAuthSettings } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    requireSession(request, "manage_access");
    return Response.json({ settings: await getOAuthSettings() }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export async function PUT(request: Request) {
  try {
    const session = requireSession(request, "manage_access");
    const body = await request.json() as { clientId?: string; clientSecret?: string; workspaceDomain?: string };
    return Response.json(
      { settings: await saveOAuthSettings(body, session.actor) },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
