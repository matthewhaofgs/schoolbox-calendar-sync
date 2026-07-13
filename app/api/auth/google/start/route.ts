import { beginGoogleOAuth } from "@/lib/auth";
import { jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    const result = await beginGoogleOAuth(request);
    return new Response(null, {
      status: 302,
      headers: { Location: result.url, "Set-Cookie": result.cookie, "Cache-Control": "no-store" },
    });
  } catch (error) {
    return jsonError(error);
  }
}
