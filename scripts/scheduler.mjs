import { setTimeout as delay } from "node:timers/promises";

try { process.loadEnvFile?.(".env.production"); } catch {}
try { process.loadEnvFile?.(".env.local"); } catch {}
try { process.loadEnvFile?.(".env"); } catch {}

const baseUrl = (process.env.INTERNAL_APP_URL || `http://127.0.0.1:${process.env.PORT || "3000"}`).replace(/\/$/, "");
const token = process.env.SCHEDULER_TOKEN?.trim();
const configuredPollMs = Number(process.env.SCHEDULER_POLL_MS || 60_000);
const pollMs = Number.isFinite(configuredPollMs) ? Math.max(15_000, configuredPollMs) : 60_000;
let stopping = false;
let activeController = null;
let lastStatus = "";
const shutdown = new AbortController();

if (!token || token.length < 32) {
  process.stderr.write("Relay scheduler: SCHEDULER_TOKEN must contain at least 32 random characters.\n");
  process.exit(1);
}

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => {
    stopping = true;
    activeController?.abort();
    shutdown.abort();
  });
}

async function tick() {
  const controller = new AbortController();
  activeController = controller;
  const timeout = setTimeout(() => controller.abort(), 2 * 60 * 60 * 1000);
  try {
    const response = await fetch(`${baseUrl}/api/sync/local-tick`, {
      method: "POST",
      headers: { Authorization: `Bearer ${token}` },
      signal: controller.signal,
    });
    const body = await response.text();
    if (!response.ok) throw new Error(`${response.status} ${body.slice(0, 300)}`);
    let status = "unknown";
    try { status = JSON.parse(body).status || status; } catch {}
    if (status !== lastStatus || status === "started") process.stdout.write(`Relay scheduler ${new Date().toISOString()}: ${body}\n`);
    lastStatus = status;
  } finally {
    clearTimeout(timeout);
    activeController = null;
  }
}

async function pause(milliseconds) {
  try { await delay(milliseconds, undefined, { signal: shutdown.signal }); } catch {}
}

while (!stopping) {
  try {
    await tick();
    await pause(pollMs);
  } catch (error) {
    if (stopping) break;
    process.stderr.write(`Relay scheduler ${new Date().toISOString()}: ${error instanceof Error ? error.message : String(error)}\n`);
    await pause(5_000);
  }
}
