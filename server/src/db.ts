/**
 * Persistent state in one SQLite database, DATA_DIR/cloudshare.db, through Node's built-in node:sqlite (Node 24) –
 * no extra package, no database server. One table holds one row per entry of each store: password jobs, password
 * entries for sending again, delivery states, download/upload notices, the cleanup list and the Nextcloud connections.
 *
 * Every row is sealed with CREDENTIAL_KEY (AES-256-GCM; the AAD is store and key, so a row cannot be moved to another
 * entry), so the file holds no readable data – as the JSON files before it. Each module keeps its entries in memory
 * (PersistentMap) and writes only the row that changed, instead of the whole file on every change.
 */
import { chmodSync, closeSync, openSync } from "node:fs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
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
  CREATE TABLE IF NOT EXISTS entries (
    store TEXT NOT NULL,
    key   TEXT NOT NULL,
    data  TEXT NOT NULL,
    PRIMARY KEY (store, key)
  ) WITHOUT ROWID;
`);

const upsert = db.prepare(
  "INSERT INTO entries (store, key, data) VALUES (?, ?, ?) ON CONFLICT (store, key) DO UPDATE SET data = excluded.data",
);
const remove = db.prepare("DELETE FROM entries WHERE store = ? AND key = ?");
const select = db.prepare("SELECT key, data FROM entries WHERE store = ?");

const aad = (store: string, key: string) => `${store}:${key}`;

/**
 * A Map that keeps its entries in the database: loaded once at start, `set` and `delete` write the row right away.
 * An entry changed in place (`map.get(key)!.field = …`) is written with `save(key)`. A row that cannot be decrypted
 * (e.g. CREDENTIAL_KEY changed) is skipped with a warning instead of stopping the start.
 */
export class PersistentMap<V> extends Map<string, V> {
  private readonly store: string;

  constructor(store: string) {
    super();
    this.store = store;
    for (const row of select.all(store) as Array<{ key: string; data: string }>) {
      try {
        super.set(row.key, JSON.parse(unseal(JSON.parse(row.data) as Sealed, aad(store, row.key))) as V);
      } catch (err) {
        console.warn(`${store}: entry ${row.key} could not be read and is ignored: ${(err as Error).message}`);
      }
    }
  }

  override set(key: string, value: V): this {
    super.set(key, value);
    this.write(key, value);
    return this;
  }

  override delete(key: string): boolean {
    const had = super.delete(key);
    if (had) remove.run(this.store, key);
    return had;
  }

  /** Writes an entry that was changed in place. */
  save(key: string): void {
    const value = super.get(key);
    if (value !== undefined) this.write(key, value);
  }

  private write(key: string, value: V) {
    upsert.run(this.store, key, JSON.stringify(seal(JSON.stringify(value), aad(this.store, key))));
  }
}
