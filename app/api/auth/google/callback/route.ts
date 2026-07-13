import { applicationOrigin, completeGoogleOAuth } from "@/lib/auth";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    const result = await completeGoogleOAuth(request);
    return new Response(null, {
      status: 302,
      headers: { Location: applicationOrigin(), "Set-Cookie": result.cookie, "Cache-Control": "no-store" },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Google sign-in failed";
    const destination = new URL(applicationOrigin());
    destination.searchParams.set("authError", message);
    return new Response(null, { status: 302, headers: { Location: destination.toString(), "Cache-Control": "no-store" } });
  }
}
