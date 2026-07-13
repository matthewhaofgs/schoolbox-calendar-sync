import { authReadiness, currentSession } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    const session = currentSession(request);
    const readiness = authReadiness();
    return Response.json(
      session ? { authenticated: true, session, readiness } : { authenticated: false, readiness },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
