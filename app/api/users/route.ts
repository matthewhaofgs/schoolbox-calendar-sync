import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import { listUserMappings, setUsersSyncEnabled } from "@/lib/storage";
import { cleanupUserManagedEvents } from "@/lib/sync";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const rawLimit = new URL(request.url).searchParams.get("limit");
    let limit: number | undefined;
    if (rawLimit !== null) {
      limit = Number(rawLimit);
      if (!Number.isInteger(limit) || limit < 1) throw new HttpError(400, "The user limit must be a positive integer");
    }
    return Response.json({ users: await listUserMappings(limit) });
  } catch (error) {
    return jsonError(error);
  }
}

export async function PATCH(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const body = await request.json() as { userIds?: unknown; syncEnabled?: unknown };
    if (!Array.isArray(body.userIds) || !body.userIds.every((id) => typeof id === "string")) {
      throw new HttpError(400, "User IDs must be supplied as a list");
    }
    if (typeof body.syncEnabled !== "boolean") throw new HttpError(400, "Choose whether these users should sync");
    const updated = await setUsersSyncEnabled(body.userIds, body.syncEnabled, actor);
    return Response.json({ updated, syncEnabled: body.syncEnabled });
  } catch (error) {
    return jsonError(error);
  }
}

export async function DELETE(request: Request) {
  try {
    const actor = await requestActor(request, "configure");
    const body = await request.json() as { userId?: unknown };
    if (typeof body.userId !== "string" || !body.userId.trim()) {
      throw new HttpError(400, "Choose a user to clean up");
    }
    return Response.json(await cleanupUserManagedEvents(body.userId, actor));
  } catch (error) {
    return jsonError(error);
  }
}
