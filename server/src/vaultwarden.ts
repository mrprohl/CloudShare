import { HttpError } from "./auth.js";
import { config } from "./config.js";

/** Talks to the local Vault Management API of `bw serve` (127.0.0.1 only, service account). */
async function bw<T>(method: string, path: string, body?: unknown): Promise<T> {
  const res = await fetch(config.bwServeUrl + path, {
    method,
    headers: { "Content-Type": "application/json" },
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const json = (await res.json().catch(() => null)) as { success?: boolean; message?: string; data?: T } | null;
  if (!res.ok || !json?.success) {
    throw new Error(json?.message || `bw serve: HTTP ${res.status}`);
  }
  return json.data as T;
}

export async function unlockVault(): Promise<void> {
  await bw("POST", "/unlock", { password: config.bwPassword });
}

/** Waits at startup until `bw serve` is reachable and unlocks the vault. */
export async function initVault(log: (msg: string) => void, attempts = 30): Promise<void> {
  for (let i = 1; ; i++) {
    try {
      await unlockVault();
      log("Vaultwarden vault unlocked");
      return;
    } catch (err) {
      if (i >= attempts) throw err;
      await new Promise((r) => setTimeout(r, 1000));
    }
  }
}

/** Runs a vault call; if it fails (vault locked, e.g. after bw serve restarted), unlocks once and retries. */
async function withUnlock<T>(call: () => Promise<T>): Promise<T> {
  try {
    return await call();
  } catch {
    try {
      await unlockVault();
      return await call();
    } catch (err) {
      throw new HttpError(502, `Vaultwarden-Send fehlgeschlagen: ${(err as Error).message}`);
    }
  }
}

/** Changes how often an existing Send can be opened (until the mail is sent, "Kennwort abrufbar"). */
export async function setSendMaxAccess(id: string, maxAccessCount: number): Promise<void> {
  const path = `/object/send/${encodeURIComponent(id)}`;
  await withUnlock(async () => {
    const send = await bw<Record<string, unknown>>("GET", path);
    await bw("PUT", path, { ...send, maxAccessCount });
  });
}

/**
 * Creates a hidden text Send containing the password (opened at most `maxAccessCount` times) and returns its URL
 * and ID (the ID allows changing the number of accesses later).
 */
export async function createPasswordSend(
  label: string,
  password: string,
  maxAccessCount: number = config.sendMaxAccessCount,
): Promise<{ url: string; id: string | undefined }> {
  const deletionDate = new Date(Date.now() + config.sendExpireDays * 86_400_000).toISOString();
  const send = {
    name: label,
    notes: null,
    type: 0,
    text: { text: password, hidden: true },
    file: null,
    maxAccessCount,
    deletionDate,
    expirationDate: deletionDate,
    password: null,
    disabled: false,
    hideEmail: true,
  };

  const data = await withUnlock(() => bw<{ accessUrl?: string; id?: string }>("POST", "/object/send", send));
  if (!data?.accessUrl) throw new HttpError(502, "Vaultwarden: Send-URL fehlt in der Antwort");
  return { url: data.accessUrl, id: typeof data.id === "string" ? data.id : undefined };
}
