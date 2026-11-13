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

/** Creates a hidden text Send containing the password (opened at most `maxAccessCount` times) and returns its URL. */
export async function createPasswordSend(
  label: string,
  password: string,
  maxAccessCount: number = config.sendMaxAccessCount,
): Promise<string> {
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

  const create = () => bw<{ accessUrl?: string }>("POST", "/object/send", send);
  let data;
  try {
    data = await create();
  } catch {
    // Vault may be locked (e.g. after bw serve restarted) – unlock once and retry.
    try {
      await unlockVault();
      data = await create();
    } catch (err) {
      throw new HttpError(502, `Vaultwarden-Send fehlgeschlagen: ${(err as Error).message}`);
    }
  }
  if (!data?.accessUrl) throw new HttpError(502, "Vaultwarden: Send-URL fehlt in der Antwort");
  return data.accessUrl;
}
