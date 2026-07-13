import { changeLocalPassword, clearSessionCookie, requireSession } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function PUT(request: Request) {
  try {
    const session = requireSession(request, "manage_access");
    const body = await request.json() as { currentPassword?: string; nextPassword?: string };
    changeLocalPassword(session, body.currentPassword ?? "", body.nextPassword ?? "");
    return Response.json(
      { changed: true, message: "Password changed. Sign in again with the new password." },
      { headers: { "Set-Cookie": clearSessionCookie(), "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
