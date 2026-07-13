import Database from "better-sqlite3";
import { pbkdf2Sync, randomBytes } from "node:crypto";
import { chmodSync } from "node:fs";
import { resolve } from "node:path";
import { createInterface } from "node:readline/promises";
import { stdin, stdout } from "node:process";

try { process.loadEnvFile?.(".env.production"); } catch {}
try { process.loadEnvFile?.(".env.local"); } catch {}
try { process.loadEnvFile?.(".env"); } catch {}
process.umask?.(0o077);

async function readSecret(label) {
  if (!stdin.isTTY || !stdin.setRawMode) throw new Error("Run this command in an interactive server terminal.");
  stdout.write(label);
  stdin.setRawMode(true);
  stdin.resume();
  let value = "";
  return new Promise((resolveSecret, reject) => {
    const restoreTerminal = () => {
      stdin.off("data", onData);
      stdin.setRawMode(false);
      stdin.pause();
    };
    const onData = (buffer) => {
      for (const character of buffer.toString("utf8")) {
        if (character === "\r" || character === "\n") {
          restoreTerminal();
          stdout.write("\n");
          resolveSecret(value);
          return;
        }
        if (character === "\u0003") {
          restoreTerminal();
          stdout.write("\n");
          reject(new Error("Cancelled"));
          return;
        }
        if (character === "\u007f" || character === "\b") {
          if (value) {
            value = value.slice(0, -1);
            stdout.write("\b \b");
          }
          continue;
        }
        if (character >= " ") {
          value += character;
          stdout.write("*");
        }
      }
    };
    stdin.on("data", onData);
  });
}

const databasePath = resolve(process.env.DATABASE_PATH || "data/relay.sqlite");
const database = new Database(databasePath);
database.pragma("busy_timeout = 5000");
database.pragma("journal_mode = WAL");
database.pragma("foreign_keys = ON");

try {
  const owners = database.prepare(
    "SELECT id, username FROM auth_users WHERE provider = 'local' AND is_owner = 1 ORDER BY created_at",
  ).all();
  if (owners.length === 0) throw new Error("No local administrator exists. Run npm run auth:bootstrap first.");
  if (owners.length !== 1) throw new Error("Recovery stopped because the database contains more than one local owner.");

  const owner = owners[0];
  const prompt = createInterface({ input: stdin, output: stdout });
  const confirmation = (await prompt.question(
    `This will reset '${owner.username}' and revoke every active session. Type the username to continue: `,
  )).trim().toLowerCase();
  prompt.close();
  if (confirmation !== owner.username.toLowerCase()) throw new Error("The username did not match. No changes were made.");

  const password = await readSecret("New password (14+ characters): ");
  const repeated = await readSecret("Confirm new password: ");
  if (password.length < 14) throw new Error("The administrator password must be at least 14 characters.");
  if (password !== repeated) throw new Error("The passwords do not match.");

  const salt = randomBytes(16);
  const iterations = 600_000;
  const digest = pbkdf2Sync(password, salt, iterations, 32, "sha256").toString("base64url");
  const hash = `pbkdf2-sha256$${iterations}$${salt.toString("base64url")}$${digest}`;
  const now = new Date().toISOString();

  database.exec(`CREATE TABLE IF NOT EXISTS auth_sessions (
    token_hash TEXT PRIMARY KEY,
    user_id TEXT NOT NULL REFERENCES auth_users(id) ON DELETE CASCADE,
    created_at TEXT NOT NULL,
    last_seen_at TEXT NOT NULL,
    absolute_expires_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    occurred_at TEXT NOT NULL,
    actor TEXT NOT NULL,
    action TEXT NOT NULL,
    detail TEXT
  )`);

  database.transaction(() => {
    database.prepare(`UPDATE auth_users
      SET password_hash = ?, failed_attempts = 0, locked_until = NULL, updated_at = ?
      WHERE id = ?`).run(hash, now, owner.id);
    database.prepare("DELETE FROM auth_sessions").run();
    database.prepare(
      "INSERT INTO audit_log (occurred_at, actor, action, detail) VALUES (?, 'server-console', 'authentication.password_reset', ?)",
    ).run(now, `Local administrator '${owner.username}' was recovered from the server console; all sessions were revoked.`);
  })();

  try { chmodSync(databasePath, 0o600); } catch {}
  stdout.write(`Password reset for '${owner.username}'. All Relay sessions have been revoked.\n`);
} finally {
  database.close();
}
