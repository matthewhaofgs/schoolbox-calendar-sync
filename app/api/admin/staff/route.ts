import { deleteStaffAccount, listStaffAccounts, requireSession, saveStaffAccount } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";

export const dynamic = "force-dynamic";

export async function GET(request: Request) {
  try {
    requireSession(request, "manage_access");
    return Response.json({ staff: listStaffAccounts() }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export async function PUT(request: Request) {
  try {
    const session = requireSession(request, "manage_access");
    const body = await request.json() as { id?: string; email?: string; displayName?: string; role?: string; enabled?: boolean };
    return Response.json({ staff: saveStaffAccount(body, session.actor) }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export async function DELETE(request: Request) {
  try {
    const session = requireSession(request, "manage_access");
    const id = new URL(request.url).searchParams.get("id");
    if (!id) throw new HttpError(400, "A staff account ID is required");
    deleteStaffAccount(id, session.actor);
    return Response.json({ deleted: true }, { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}
