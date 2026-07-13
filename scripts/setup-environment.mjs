import { randomBytes } from "node:crypto";
import { existsSync, writeFileSync } from "node:fs";
import { createInterface } from "node:readline/promises";
import { stdin, stdout } from "node:process";

const development = process.argv.includes("--development");
const target = development ? ".env.local" : ".env.production";
if (existsSync(target)) {
  process.stderr.write(`${target} already exists. Move or remove it before generating a replacement.\n`);
  process.exit(1);
}

const prompt = createInterface({ input: stdin, output: stdout });
const entered = development ? "http://127.0.0.1:3000" : (await prompt.question("Internal HTTPS origin [https://relay.internal.example]: ")).trim();
prompt.close();
const appOrigin = (entered || "https://relay.internal.example").replace(/\/$/, "");
let parsed;
try { parsed = new URL(appOrigin); } catch { process.stderr.write("Enter a valid absolute URL.\n"); process.exit(1); }
if ((!development && parsed.protocol !== "https:") || parsed.origin !== appOrigin) {
  process.stderr.write("Production APP_ORIGIN must be an HTTPS origin without a path.\n");
  process.exit(1);
}

const encoded = (size) => randomBytes(size).toString("base64");
const contents = [
  `NODE_ENV=${development ? "development" : "production"}`,
  "RELAY_HOST=0.0.0.0",
  "PORT=3000",
  `APP_ORIGIN=${appOrigin}`,
  "DATABASE_PATH=data/relay.sqlite",
  `CONFIG_ENCRYPTION_KEY=${encoded(32)}`,
  `SESSION_SECRET=${encoded(48)}`,
  `SCHEDULER_TOKEN=${encoded(48)}`,
  "",
].join("\n");
writeFileSync(target, contents, { encoding: "utf8", mode: 0o600, flag: "wx" });
stdout.write(`Created ${target}. Keep it private and back it up with the database.\n`);
