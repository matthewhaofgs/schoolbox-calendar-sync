import Database from "better-sqlite3";
import { mkdirSync } from "node:fs";
import { dirname, isAbsolute, join } from "node:path";

type BoundValue = string | number | bigint | null | Uint8Array;

class BoundStatement {
  private values: BoundValue[] = [];

  constructor(private readonly statement: Database.Statement) {}

  bind(...values: BoundValue[]): this {
    this.values = values;
    return this;
  }

  run() {
    return this.statement.run(...this.values);
  }

  first<T>(): T | null {
    return (this.statement.get(...this.values) as T | undefined) ?? null;
  }

  all<T>(): { results: T[] } {
    return { results: this.statement.all(...this.values) as T[] };
  }
}

export class RelayDatabase {
  constructor(private readonly native: Database.Database) {}

  prepare(sql: string): BoundStatement {
    return new BoundStatement(this.native.prepare(sql));
  }

  batch(statements: BoundStatement[]): void {
    this.native.transaction(() => {
      for (const statement of statements) statement.run();
    })();
  }

  transaction<T>(operation: () => T): T {
    return this.native.transaction(operation)();
  }

  pragma(source: string): unknown {
    return this.native.pragma(source);
  }

  close(): void {
    this.native.close();
  }
}

const globalDatabase = globalThis as typeof globalThis & {
  __relayDatabase?: RelayDatabase;
};

export function databasePath(): string {
  const configured = process.env.DATABASE_PATH?.trim();
  if (!configured) return join(process.cwd(), "data", "relay.sqlite");
  if (isAbsolute(configured)) return configured;
  return join(/* turbopackIgnore: true */ process.cwd(), configured);
}

export function db(): RelayDatabase {
  if (globalDatabase.__relayDatabase) return globalDatabase.__relayDatabase;

  const path = databasePath();
  mkdirSync(dirname(path), { recursive: true });
  const native = new Database(path);
  native.pragma("journal_mode = WAL");
  native.pragma("foreign_keys = ON");
  native.pragma("busy_timeout = 5000");
  native.pragma("synchronous = NORMAL");

  globalDatabase.__relayDatabase = new RelayDatabase(native);
  return globalDatabase.__relayDatabase;
}
