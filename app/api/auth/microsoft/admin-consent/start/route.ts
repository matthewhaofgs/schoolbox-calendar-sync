import { applicationOrigin, requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import { createMicrosoftConsentState, getStoredMicrosoftConnection } from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function POST(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const connection = await getStoredMicrosoftConnection();
    if (!connection.tenantId || !connection.clientId || !connection.clientSecret) {
      throw new HttpError(409, "Save the Microsoft tenant ID, application client ID, and client secret before requesting admin consent");
    }
    if (!connection.testUserEmail) {
      throw new HttpError(409, "Choose an explicit Microsoft 365 test mailbox before requesting admin consent");
    }
    const state = await createMicrosoftConsentState(connection, actor);
    const callbackUrl = `${applicationOrigin()}/api/auth/microsoft/admin-consent/callback`;
    const destination = new URL(`https://login.microsoftonline.com/${encodeURIComponent(connection.tenantId)}/v2.0/adminconsent`);
    destination.searchParams.set("client_id", connection.clientId);
    destination.searchParams.set("scope", "https://graph.microsoft.com/.default");
    destination.searchParams.set("redirect_uri", callbackUrl);
    destination.searchParams.set("state", state);
    return Response.json(
      { url: destination.toString() },
      { headers: { "Cache-Control": "no-store" } },
    );
  } catch (error) {
    return jsonError(error);
  }
}
