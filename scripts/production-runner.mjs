import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, isAbsolute, join, resolve } from "node:path";

try { process.loadEnvFile?.(".env.production"); } catch {}
try { process.loadEnvFile?.(".env.local"); } catch {}
try { process.loadEnvFile?.(".env"); } catch {}
process.umask?.(0o077);

const root = process.cwd();
const configuredDatabasePath = process.env.DATABASE_PATH || "data/relay.sqlite";
const environment = {
  ...process.env,
  NODE_ENV: "production",
  HOSTNAME: process.env.RELAY_HOST || "0.0.0.0",
  PORT: process.env.PORT || "3000",
  DATABASE_PATH: isAbsolute(configuredDatabasePath) ? configuredDatabasePath : resolve(root, configuredDatabasePath),
};

const rootServer = join(root, "server.js");
const builtServer = join(root, ".next", "standalone", "server.js");
const serverEntry = existsSync(rootServer) ? rootServer : existsSync(builtServer) ? builtServer : null;
const serverDirectory = serverEntry ? dirname(serverEntry) : root;
const webArgs = serverEntry
  ? [serverEntry]
  : [join(root, "node_modules", "next", "dist", "bin", "next"), "start", "--hostname", environment.HOSTNAME, "--port", environment.PORT];

const children = [
  spawn(process.execPath, webArgs, { cwd: serverDirectory, env: environment, stdio: "inherit" }),
  spawn(process.execPath, [join(root, "scripts", "scheduler.mjs")], { cwd: root, env: environment, stdio: "inherit" }),
];

let exiting = false;
let exitCode = 0;
const remaining = new Set(children);
function stop(signal = "SIGTERM", code = 0) {
  if (exiting) return;
  exiting = true;
  exitCode = code;
  for (const child of children) {
    if (!child.killed) child.kill(signal);
  }
  setTimeout(() => {
    for (const child of remaining) {
      child.kill("SIGKILL");
    }
    process.exit(exitCode);
  }, 25_000).unref();
}

for (const signal of ["SIGINT", "SIGTERM"]) process.on(signal, () => stop(signal, 0));
for (const child of children) {
  child.on("exit", (code, signal) => {
    remaining.delete(child);
    if (!exiting) stop(signal || "SIGTERM", code ?? 1);
    if (exiting && remaining.size === 0) process.exit(exitCode);
  });
  child.on("error", (error) => {
    process.stderr.write(`Relay process failed: ${error.message}\n`);
    stop("SIGTERM", 1);
  });
}
