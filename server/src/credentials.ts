/**
 * Persistent storage of the Nextcloud app passwords per Entra user (oid), in the database (db.ts, store
 * "credentials"): every entry sealed with CREDENTIAL_KEY, bound to the user's oid.
 */
import { PersistentMap } from "./db.js";

export interface NextcloudCredentials {
  loginName: string;
  appPassword: string;
}

interface StoredEntry extends NextcloudCredentials {
  connectedAt: string;
}

const entries = new PersistentMap<StoredEntry>("credentials");

export function getCredentials(oid: string): NextcloudCredentials | undefined {
  const entry = entries.get(oid);
  return entry ? { loginName: entry.loginName, appPassword: entry.appPassword } : undefined;
}

export async function saveCredentials(oid: string, creds: NextcloudCredentials): Promise<void> {
  entries.set(oid, { loginName: creds.loginName, appPassword: creds.appPassword, connectedAt: new Date().toISOString() });
}

export async function removeCredentials(oid: string): Promise<void> {
  entries.delete(oid);
}
