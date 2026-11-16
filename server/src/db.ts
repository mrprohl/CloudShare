/**
 * Persistent state in one SQLite database, DATA_DIR/cloudshare.db, through Node's built-in node:sqlite (Node 24) –
 * no extra package, no database server. One table holds one row per entry of each store: password jobs, password
 * entries for sending again, delivery states, download/upload notices, the cleanup list and the Nextcloud connections.
 *
 * Every row is sealed with CREDENTIAL_KEY (AES-256-GCM; the AAD is store and key, so a row cannot be moved to another
 * entry), so the file holds no readable data – as the JSON files before it. Each module keeps its entries in memory
 * (PersistentMap) and writes only the row that changed, instead of the whole file on every change.
 *
 * Entries leave the database on their own: jobs 7 days after their last link, password entries and notices when the
 * links expire, delivery states after 120 days, cleanup entries when the folder was deleted, connections on
 * disconnect. As a safety net a daily pass deletes every entry not changed for DATA_RETENTION_DAYS (default 365) –
 * except the Nextcloud connections (in use as long as the user is) and the cleanup list (removed when due) – and then
 * gives the freed space back to the file system (VACUUM).
 */
import { chmodSync, closeSync, openSync } from "node:fs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
import type { FastifyBaseLogger } from "fastify";
import { config } from "./config.js";
import { seal, unseal, type Sealed } from "./storage.js";

const file = path.join(config.dataDir, "cloudshare.db");
// Readable only by the CloudShare user (600), like the files before it; SQLite gives its -wal/-shm files the same mode.
closeSync(openSync(file, "a", 0o600));
chmodSync(file, 0o600);
const db = new DatabaseSync(file);
db.exec(`
  PRAGMA journal_mode = WAL;
  PRAGMA synchronous = NORMAL;
  PRAGMA busy_timeout = 5000;
  PRAGMA temp_store = MEMORY;
  CREATE TABLE IF NOT EXISTS entries (
    store   TEXT NOT NULL,
    key     TEXT NOT NULL,
    data    TEXT NOT NULL,
    updated INTEGER NOT NULL DEFAULT 0,
    PRIMARY KEY (store, key)
  ) WITHOUT ROWID;
`);
// Databases created before the column existed get it (their rows count as changed at the first start with it).
const columns = db.prepare("PRAGMA table_info(entries)").all() as Array<{ name: string }>;
if (!columns.some((c) => c.name === "updated")) db.exec("ALTER TABLE entries ADD COLUMN updated INTEGER NOT NULL DEFAULT 0");

const upsert = db.prepare(
  "INSERT INTO entries (store, key, data, updated) VALUES (?, ?, ?, ?) " +
    "ON CONFLICT (store, key) DO UPDATE SET data = excluded.data, updated = excluded.updated",
);
const remove = db.prepare("DELETE FROM entries WHERE store = ? AND key = ?");
const select = db.prepare("SELECT key, data, updated FROM entries WHERE store = ?");

const aad = (store: string, key: string) => `${store}:${key}`;

/** All maps the retention pass looks at (those created with `retention: true`). */
const retained: Array<PersistentMap<unknown>> = [];

/**
 * A Map that keeps its entries in the database: loaded once at start, `set` and `delete` write the row right away.
 * An entry changed in place (`map.get(key)!.field = …`) is written with `save(key)`. A row that cannot be decrypted
 * (e.g. CREDENTIAL_KEY changed) is skipped with a warning instead of stopping the start.
 */
export class PersistentMap<V> extends Map<string, V> {
  private readonly store: string;
  /** When each entry was last written (ms) – for DATA_RETENTION_DAYS. */
  private readonly updated = new Map<string, number>();

  /** `retention: false` keeps the entries however old they are (see the module comment). */
  constructor(store: string, options: { retention?: boolean } = {}) {
    super();
    this.store = store;
    const now = Date.now();
    for (const row of select.all(store) as Array<{ key: string; data: string; updated: number }>) {
      try {
        super.set(row.key, JSON.parse(unseal(JSON.parse(row.data) as Sealed, aad(store, row.key))) as V);
        this.updated.set(row.key, row.updated || now);
      } catch (err) {
        console.warn(`${store}: entry ${row.key} could not be read and is ignored: ${(err as Error).message}`);
      }
    }
    if (options.retention !== false) retained.push(this as PersistentMap<unknown>);
  }

  override set(key: string, value: V): this {
    super.set(key, value);
    this.write(key, value);
    return this;
  }

  override delete(key: string): boolean {
    const had = super.delete(key);
    this.updated.delete(key);
    if (had) remove.run(this.store, key);
    return had;
  }

  /** Writes an entry that was changed in place. */
  save(key: string): void {
    const value = super.get(key);
    if (value !== undefined) this.write(key, value);
  }

  /** Deletes the entries last written before `cutoff` (ms); returns how many. */
  prune(cutoff: number): number {
    let removed = 0;
    for (const [key, at] of [...this.updated]) {
      if (at < cutoff && this.delete(key)) removed++;
    }
    return removed;
  }

  private write(key: string, value: V) {
    const now = Date.now();
    upsert.run(this.store, key, JSON.stringify(seal(JSON.stringify(value), aad(this.store, key))), now);
    this.updated.set(key, now);
  }
}

/**
 * One retention pass (exported for tests): deletes the entries not changed for `days` days from every map with
 * retention and shrinks the file when something was deleted. Returns the number of deleted entries.
 */
export function pruneOldEntries(days: number, now = Date.now()): number {
  const cutoff = now - days * 86_400_000;
  const removed = retained.reduce((sum, map) => sum + map.prune(cutoff), 0);
  if (removed > 0) db.exec("VACUUM");
  return removed;
}

/** Daily retention pass (DATA_RETENTION_DAYS); the first one a few minutes after start. */
export function startRetention(log: Pick<FastifyBaseLogger, "info" | "warn">) {
  const days = config.dataRetentionDays;
  if (days === null) {
    log.info("Retention of stored entries is off (DATA_RETENTION_DAYS=off)");
    return;
  }
  const pass = () => {
    try {
      const removed = pruneOldEntries(days);
      if (removed > 0) log.info({ removed, days }, "Old entries removed from the database");
    } catch (err) {
      log.warn({ err }, "Retention pass failed, retrying tomorrow");
    }
  };
  setTimeout(pass, 5 * 60_000).unref();
  setInterval(pass, 24 * 60 * 60_000).unref();
}
