import { GoogleWorkspaceClient, parseServiceAccountJson } from "@/lib/google";
import { MicrosoftGraphClient } from "@/lib/microsoft";
import { SchoolboxClient } from "@/lib/schoolbox";
import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import {
  getStoredGoogleConnection,
  getStoredMicrosoftConnection,
  getStoredSchoolboxConnection,
  listRuns,
  recordConnectionVerified,
  recordMicrosoftAdminConsent,
} from "@/lib/storage";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    return Response.json({ runs: await listRuns(10) });
  } catch (error) {
    return jsonError(error);
  }
}

type DiagnosticPayload = {
  target?: "schoolbox" | "google" | "microsoft";
  config?: {
    schoolboxUrl?: string;
    schoolboxJwt?: string;
    serviceAccountJson?: string;
    adminEmail?: string;
    googleCustomer?: string;
    microsoftTenantId?: string;
    microsoftClientId?: string;
    microsoftClientSecret?: string;
    microsoftTestUserEmail?: string;
  };
};

export async function POST(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const body = (await request.json()) as DiagnosticPayload;

    if (body.target === "schoolbox") {
      const stored = await getStoredSchoolboxConnection();
      const baseUrl = body.config?.schoolboxUrl?.trim() || stored.baseUrl;
      const suppliedJwt = body.config?.schoolboxJwt?.trim();
      if (baseUrl && stored.baseUrl && !suppliedJwt) {
        let originChanged = false;
        try { originChanged = new URL(baseUrl).origin !== new URL(stored.baseUrl).origin; } catch { originChanged = true; }
        if (originChanged) throw new Error("Enter a new Schoolbox JWT when testing a different Schoolbox host.");
      }
      const jwt = suppliedJwt || stored.token;
      if (!baseUrl || !jwt) throw new Error("Enter the Schoolbox base URL and superuser JWT before testing.");
      const result = await new SchoolboxClient({ baseUrl, jwt, pastDays: 1, futureDays: 7 }).testConnection();
      if (!body.config) await recordConnectionVerified("schoolbox", actor, stored.credentialVersion);
      return Response.json({
        ok: true,
        target: "schoolbox",
        message: `Schoolbox verified: ${result.usersVisible} users sampled and delegated calendar access confirmed.`,
        result,
      });
    }

    if (body.target === "google") {
      const stored = await getStoredGoogleConnection();
      const credentialJson = body.config?.serviceAccountJson?.trim() || stored.serviceAccountJson;
      const adminEmail = body.config?.adminEmail?.trim() || stored.adminEmail;
      const customer = body.config?.googleCustomer?.trim() || stored.customer;
      if (!credentialJson || !adminEmail) throw new Error("Add the service-account JSON and delegated admin email before testing.");
      const result = await new GoogleWorkspaceClient(parseServiceAccountJson(credentialJson)).testConnection({
        adminSubject: adminEmail,
        customer,
      });
      if (!body.config) await recordConnectionVerified("google", actor, stored.credentialVersion);
      return Response.json({
        ok: true,
        target: "google",
        message: `Google Workspace verified for ${result.directory.adminSubject}; event access and app-created calendar delegation both succeeded.`,
        result,
      });
    }

    if (body.target === "microsoft") {
      const stored = await getStoredMicrosoftConnection();
      const tenantId = (body.config?.microsoftTenantId?.trim() || stored.tenantId).toLowerCase();
      const clientId = (body.config?.microsoftClientId?.trim() || stored.clientId).toLowerCase();
      const clientSecret = body.config?.microsoftClientSecret?.trim() || stored.clientSecret;
      const testUserEmail = body.config?.microsoftTestUserEmail?.trim() || stored.testUserEmail;
      if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Add the Microsoft tenant ID, client ID, and client credential before testing.");
      }
      if (!testUserEmail) {
        throw new Error("Choose an explicit Microsoft 365 test mailbox before testing calendar write access.");
      }
      const result = await new MicrosoftGraphClient({ tenantId, clientId, clientSecret }).testConnection({
        targetUserId: testUserEmail,
      });
      // A successful app-only Graph probe proves that tenant admin consent is
      // effective even when it was granted directly in the Entra portal.
      const savedCredentials = !body.config;
      const microsoftConsentGrantedAt = savedCredentials
        ? await recordMicrosoftAdminConsent(actor, stored)
        : undefined;
      return Response.json({
        ok: true,
        target: "microsoft",
        message: "Microsoft 365 verified: Entra discovery, Outlook primary-calendar access, and temporary secondary-calendar create/delete succeeded.",
        microsoftConsentGrantedAt,
        result,
      });
    }

    throw new Error("Choose Schoolbox, Google Workspace, or Microsoft 365 diagnostics.");
  } catch (error) {
    return jsonError(error);
  }
}
