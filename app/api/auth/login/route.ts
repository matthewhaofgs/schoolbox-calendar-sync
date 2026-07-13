import { assertRequestOrigin, localLogin } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    assertRequestOrigin(request);
    const length = Number(request.headers.get("content-length") ?? 0);
    if (length > 16_384) throw new HttpError(413, "Login request is too large");
    const body = await request.json() as { username?: string; password?: string };
    const result = localLogin(body.username ?? "", body.password ?? "");
    return Response.json(
      { authenticated: true, session: result.session },
      { headers: { "Set-Cookie": result.cookie, "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
