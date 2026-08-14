import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import { getConfig, saveConfig, type ConfigInput } from "@/lib/storage";

export const dynamic = "force-dynamic";

const stringFields = [
  "schoolboxBaseUrl", "schoolboxToken", "googleServiceAccountJson", "googleAdminEmail",
  "googleCustomer", "timezone", "microsoftTenantId", "microsoftClientId",
  "microsoftClientSecret", "microsoftTestUserEmail",
] as const satisfies readonly (keyof ConfigInput)[];
const numberFields = [
  "pastDays", "futureDays", "concurrency", "discoveryTimeoutSeconds",
  "userSyncTimeoutSeconds", "runTimeoutMinutes", "syncIntervalMinutes",
] as const satisfies readonly (keyof ConfigInput)[];
const booleanFields = [
  "syncNewUsersByDefault", "googleEnabled", "syncNewGoogleUsersByDefault",
  "microsoftEnabled", "syncNewMicrosoftUsersByDefault", "enabled",
  "schoolboxSetupCompleted", "googleSetupCompleted", "microsoftSetupCompleted",
] as const satisfies readonly (keyof ConfigInput)[];

function validatedConfigInput(value: unknown): ConfigInput {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new HttpError(400, "Configuration must be a JSON object");
  }
  const input = value as Record<string, unknown>;
  const allowed = new Set<string>([...stringFields, ...numberFields, ...booleanFields, "syncPolicy", "microsoftSyncPolicy"]);
  const unknown = Object.keys(input).find((field) => !allowed.has(field));
  if (unknown) throw new HttpError(400, `${unknown} is not a writable configuration field`);
  for (const field of stringFields) {
    if (input[field] !== undefined && typeof input[field] !== "string") {
      throw new HttpError(400, `${field} must be text`);
    }
  }
  for (const field of numberFields) {
    if (input[field] !== undefined && (typeof input[field] !== "number" || !Number.isFinite(input[field]))) {
      throw new HttpError(400, `${field} must be a finite number`);
    }
  }
  for (const field of booleanFields) {
    if (input[field] !== undefined && typeof input[field] !== "boolean") {
      throw new HttpError(400, `${field} must be true or false`);
    }
  }
  for (const field of ["syncPolicy", "microsoftSyncPolicy"] as const) {
    if (input[field] !== undefined && (!input[field] || typeof input[field] !== "object" || Array.isArray(input[field]))) {
      throw new HttpError(400, `${field} must be a JSON object`);
    }
  }
  return input as ConfigInput;
}

export async function GET(request: Request) {
  try {
    await requestActor(request, "configure");
    return Response.json(await getConfig(false), { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export async function POST(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    let body: unknown;
    try { body = await request.json(); }
    catch { throw new HttpError(400, "Configuration request is not valid JSON"); }
    const input = validatedConfigInput(body);
    return Response.json(await saveConfig(input, actor), { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export const PUT = POST;
