import { requestActor } from "@/lib/auth";
import { HttpError, jsonError } from "@/lib/security";
import {
  getRun,
  listRunEventDiagnostics,
  listRunTargets,
  listRunUserDiagnostics,
  normalizeTargetProvider,
} from "@/lib/storage";

export const dynamic = "force-dynamic";

const privateJson = (body: unknown) => Response.json(body, {
  headers: { "Cache-Control": "private, no-store" },
});

function boundedInteger(value: string | null, fallback: number, minimum: number, maximum: number): number {
  if (value === null) return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < minimum || parsed > maximum) {
    throw new HttpError(400, `Expected an integer between ${minimum} and ${maximum}`);
  }
  return parsed;
}

export async function GET(request: Request) {
  try {
    await requestActor(request, "view");
    const url = new URL(request.url);
    const runId = url.searchParams.get("runId")?.trim();
    const userId = url.searchParams.get("userId")?.trim();
    const target = normalizeTargetProvider(url.searchParams.get("target") ?? "google");
    if (!runId || runId.length > 200) throw new HttpError(400, "Choose a valid run");

    const run = await getRun(runId);
    if (!run) throw new HttpError(404, "Run not found");
    const [users, targets] = await Promise.all([listRunUserDiagnostics(runId), listRunTargets(runId)]);
    if (!userId) return privateJson({ run, targets, users });
    if (userId.length > 200 || !users.some((user) => user.target === target && user.targetUserId === userId)) {
      throw new HttpError(404, "This user has no detailed outcome for the selected run");
    }

    const limit = boundedInteger(url.searchParams.get("limit"), 100, 1, 250);
    const offset = boundedInteger(url.searchParams.get("offset"), 0, 0, 1_000_000);
    return privateJson({
      run,
      targets,
      users,
      selectedUser: users.find((user) => user.target === target && user.targetUserId === userId),
      ...(await listRunEventDiagnostics(runId, userId, { limit, offset }, target)),
      limit,
      offset,
    });
  } catch (error) {
    return jsonError(error);
  }
}
