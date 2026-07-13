import { authReadiness, currentSession, extendSession } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    const session = currentSession(request, { touch: false });
    const readiness = authReadiness();
    return Response.json(
      session ? { authenticated: true, session, readiness } : { authenticated: false, readiness },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}

export async function POST(request: Request) {
  try {
    const result = extendSession(request);
    return Response.json(
      { authenticated: true, session: result.session },
      { headers: { "Set-Cookie": result.cookie, "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
