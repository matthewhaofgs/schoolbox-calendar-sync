import { GoogleWorkspaceClient, parseServiceAccountJson } from "@/lib/google";
import { SchoolboxClient } from "@/lib/schoolbox";
import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { getStoredGoogleConnection, getStoredSchoolboxConnection, listRuns } from "@/lib/storage";

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
  target?: "schoolbox" | "google";
  config?: {
    schoolboxUrl?: string;
    schoolboxJwt?: string;
    serviceAccountJson?: string;
    adminEmail?: string;
  };
};

export async function POST(request: Request) {
  try {
    await requestActor(request, "configure");
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
      if (!credentialJson || !adminEmail) throw new Error("Add the service-account JSON and delegated admin email before testing.");
      const result = await new GoogleWorkspaceClient(parseServiceAccountJson(credentialJson)).testConnection({
        adminSubject: adminEmail,
      });
      return Response.json({
        ok: true,
        target: "google",
        message: `Google Workspace verified for ${result.directory.adminSubject}; Calendar impersonation also succeeded.`,
        result,
      });
    }

    throw new Error("Choose either the Schoolbox or Google diagnostic.");
  } catch (error) {
    return jsonError(error);
  }
}
