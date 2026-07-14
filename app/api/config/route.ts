import { requestActor } from "@/lib/auth";
import { jsonError } from "@/lib/security";
import { getConfig, saveConfig, type ConfigInput } from "@/lib/storage";

export const dynamic = "force-dynamic";

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
    const input = (await request.json()) as ConfigInput;
    return Response.json(await saveConfig(input, actor), { headers: { "Cache-Control": "no-store" } });
  } catch (error) {
    return jsonError(error);
  }
}

export const PUT = POST;
