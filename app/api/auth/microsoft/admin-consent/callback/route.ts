import { applicationOrigin } from "@/lib/auth";
import { MicrosoftGraphClient } from "@/lib/microsoft";
import {
  consumeMicrosoftConsentState,
  getStoredMicrosoftConnection,
  recordMicrosoftAdminConsent,
} from "@/lib/storage";

export const dynamic = "force-dynamic";

function redirectWithResult(key: "microsoftConsent" | "microsoftConsentError", value: string): Response {
  const destination = new URL(applicationOrigin());
  destination.searchParams.set(key, value);
  return new Response(null, {
    status: 302,
    headers: { Location: destination.toString(), "Cache-Control": "no-store" },
  });
}

export async function GET(request: Request) {
  try {
    const params = new URL(request.url).searchParams;
    const state = params.get("state")?.trim();
    if (!state) throw new Error("Microsoft returned an incomplete admin-consent response");
    const tenant = params.get("tenant")?.trim();
    // Consume the one-time state even when Microsoft returns an error so a
    // cancelled or denied callback cannot be replayed later.
    const consent = await consumeMicrosoftConsentState(state, tenant || undefined);
    const providerError = params.get("error_description") || params.get("error");
    if (providerError) throw new Error(`Microsoft admin consent was not completed: ${providerError}`);
    if (!tenant || params.get("admin_consent")?.toLowerCase() !== "true") {
      throw new Error("Microsoft returned an incomplete admin-consent response");
    }
    const connection = await getStoredMicrosoftConnection();
    if (!connection.clientSecret || !connection.testUserEmail ||
      connection.tenantId.toLowerCase() !== consent.tenantId.toLowerCase() ||
      connection.clientId.toLowerCase() !== consent.clientId.toLowerCase() ||
      connection.credentialVersion !== consent.credentialVersion) {
      throw new Error("The saved Microsoft connection changed during admin consent");
    }
    await new MicrosoftGraphClient({
      tenantId: connection.tenantId,
      clientId: connection.clientId,
      clientSecret: connection.clientSecret,
    }).testConnection({ targetUserId: connection.testUserEmail });
    await recordMicrosoftAdminConsent(consent.actor, consent);
    return redirectWithResult("microsoftConsent", "verified");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Microsoft admin consent failed";
    return redirectWithResult("microsoftConsentError", message.slice(0, 500));
  }
}
