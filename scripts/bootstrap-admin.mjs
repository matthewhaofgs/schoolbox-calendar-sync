import Database from "better-sqlite3";
import { pbkdf2Sync, randomBytes, randomUUID } from "node:crypto";
import { chmodSync, mkdirSync } from "node:fs";
import { dirname, resolve } from "node:path";
import { createInterface } from "node:readline/promises";
import { stdin, stdout } from "node:process";

try { process.loadEnvFile?.(".env.production"); } catch {}
try { process.loadEnvFile?.(".env.local"); } catch {}
try { process.loadEnvFile?.(".env"); } catch {}
process.umask?.(0o077);

async function readSecret(label) {
  if (!stdin.isTTY || !stdin.setRawMode) throw new Error("Run this command in an interactive terminal.");
  stdout.write(label);
  stdin.setRawMode(true);
  stdin.resume();
  let value = "";
  return new Promise((resolveSecret, reject) => {
    const onData = (buffer) => {
      const text = buffer.toString("utf8");
      for (const character of text) {
        if (character === "\r" || character === "\n") {
          stdin.off("data", onData);
          stdin.setRawMode(false);
          stdin.pause();
          stdout.write("\n");
          resolveSecret(value);
          return;
        }
        if (character === "\u0003") {
          stdin.off("data", onData);
          stdin.setRawMode(false);
          reject(new Error("Cancelled"));
          return;
        }
        if (character === "\u007f" || character === "\b") {
          if (value) { value = value.slice(0, -1); stdout.write("\b \b"); }
          continue;
        }
        if (character >= " ") { value += character; stdout.write("*"); }
      }
    };
    stdin.on("data", onData);
  });
}

const databasePath = resolve(process.env.DATABASE_PATH || "data/relay.sqlite");
mkdirSync(dirname(databasePath), { recursive: true, mode: 0o700 });
const database = new Database(databasePath);
database.pragma("journal_mode = WAL");
database.pragma("foreign_keys = ON");
database.exec(`CREATE TABLE IF NOT EXISTS auth_users (
  id TEXT PRIMARY KEY,
  provider TEXT NOT NULL CHECK (provider IN ('local', 'google')),
  username TEXT UNIQUE,
  email TEXT UNIQUE,
  google_sub TEXT UNIQUE,
  display_name TEXT,
  role TEXT NOT NULL CHECK (role IN ('admin', 'operator', 'viewer')),
  is_owner INTEGER NOT NULL DEFAULT 0,
  enabled INTEGER NOT NULL DEFAULT 1,
  password_hash TEXT,
  failed_attempts INTEGER NOT NULL DEFAULT 0,
  locked_until TEXT,
  last_login_at TEXT,
  created_by TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
)`);
const existing = database.prepare("SELECT username FROM auth_users WHERE provider = 'local' AND is_owner = 1 LIMIT 1").get();
if (existing) {
  database.close();
  process.stderr.write(`A local administrator (${existing.username}) already exists. Change its password from Relay.\n`);
  process.exit(1);
}

const prompt = createInterface({ input: stdin, output: stdout });
const username = ((await prompt.question("Local administrator username [administrator]: ")).trim() || "administrator").toLowerCase();
prompt.close();
const password = await readSecret("Password (14+ characters): ");
const confirmation = await readSecret("Confirm password: ");
if (password.length < 14) throw new Error("The administrator password must be at least 14 characters.");
if (password !== confirmation) throw new Error("The passwords do not match.");

const salt = randomBytes(16);
const iterations = 600_000;
const digest = pbkdf2Sync(password, salt, iterations, 32, "sha256").toString("base64url");
const hash = `pbkdf2-sha256$${iterations}$${salt.toString("base64url")}$${digest}`;
const now = new Date().toISOString();
database.prepare(`INSERT INTO auth_users
  (id, provider, username, display_name, role, is_owner, enabled, password_hash, created_by, created_at, updated_at)
  VALUES (?, 'local', ?, 'Administrator', 'admin', 1, 1, ?, 'bootstrap', ?, ?)`)
  .run(randomUUID(), username, hash, now, now);
database.close();
try { chmodSync(databasePath, 0o600); } catch {}
stdout.write(`Created local administrator '${username}' in ${databasePath}.\n`);
