import { clearSessionCookie, logout } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    logout(request);
    return Response.json(
      { authenticated: false },
      { headers: { "Set-Cookie": clearSessionCookie(), "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
