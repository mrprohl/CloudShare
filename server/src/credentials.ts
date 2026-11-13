/**
 * Persistent storage of the Nextcloud app passwords per Entra user (oid).
 * AES-256-GCM; the oid is bound to the entry as additional authenticated data (AAD).
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { config } from "./config.js";
import { createAtomicWriter, seal, unseal } from "./storage.js";

export interface NextcloudCredentials {
  loginName: string;
  appPassword: string;
}

interface StoredEntry {
  loginName: string;
  iv: string;
  tag: string;
  secret: string;
  connectedAt: string;
}

const file = path.join(config.dataDir, "nextcloud-credentials.json");
const write = createAtomicWriter(file);

let entries: Record<string, StoredEntry> = {};
try {
  entries = JSON.parse(readFileSync(file, "utf8")) as Record<string, StoredEntry>;
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") throw err;
}

const persist = () => write(JSON.stringify(entries, null, 2));

export function getCredentials(oid: string): NextcloudCredentials | undefined {
  const entry = entries[oid];
  if (!entry) return undefined;
  try {
    const appPassword = unseal({ iv: entry.iv, tag: entry.tag, data: entry.secret }, oid);
    return { loginName: entry.loginName, appPassword };
  } catch {
    // Key changed or entry corrupted → treat as "not connected".
    return undefined;
  }
}

export async function saveCredentials(oid: string, creds: NextcloudCredentials): Promise<void> {
  const sealed = seal(creds.appPassword, oid);
  entries[oid] = {
    loginName: creds.loginName,
    iv: sealed.iv,
    tag: sealed.tag,
    secret: sealed.data,
    connectedAt: new Date().toISOString(),
  };
  await persist();
}

export async function removeCredentials(oid: string): Promise<void> {
  if (!(oid in entries)) return;
  delete entries[oid];
  await persist();
}
