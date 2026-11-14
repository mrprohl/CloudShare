import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { t } from "./i18n.js";

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
      throw new HttpError(502, t("server.vaultSendFailed", { error: (err as Error).message }));
    }
  }
}

/**
 * The password stored in a Send – for the SMS channel, right before the SMS goes out. The password is not kept
 * anywhere else on the server.
 */
export async function readSendPassword(id: string): Promise<string> {
  const send = await withUnlock(() => bw<{ text?: { text?: string } }>("GET", `/object/send/${encodeURIComponent(id)}`));
  const password = send?.text?.text;
  if (!password) throw new HttpError(502, t("server.vaultNoPassword"));
  return password;
}

/** Deletes a Send that is no longer needed (the password went out by SMS instead). */
export async function deleteSend(id: string): Promise<void> {
  await withUnlock(() => bw("DELETE", `/object/send/${encodeURIComponent(id)}`));
}

/** Moves the deletion of an existing Send to `expiresAt` (password by SMS: it must last until the mail is sent). */
export async function setSendExpiry(id: string, expiresAt: number): Promise<void> {
  const path = `/object/send/${encodeURIComponent(id)}`;
  const date = new Date(expiresAt).toISOString();
  await withUnlock(async () => {
    const send = await bw<Record<string, unknown>>("GET", path);
    await bw("PUT", path, { ...send, deletionDate: date, expirationDate: date });
  });
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
 * Creates a hidden text Send containing the password (opened at most `maxAccessCount` times, deleted after `days`)
 * and returns its URL, ID (allows changing the number of accesses later) and expiry time (ms).
 */
export async function createPasswordSend(
  label: string,
  password: string,
  maxAccessCount: number = config.sendMaxAccessCount,
  days: number = config.sendExpireDays,
): Promise<{ url: string; id: string | undefined; expiresAt: number }> {
  const expiresAt = Date.now() + days * 86_400_000;
  const deletionDate = new Date(expiresAt).toISOString();
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
  if (!data?.accessUrl) throw new HttpError(502, t("server.vaultNoUrl"));
  return { url: data.accessUrl, id: typeof data.id === "string" ? data.id : undefined, expiresAt };
}

/** Longest confidential message (characters) – Bitwarden's limit for text Sends. */
export const MESSAGE_MAX_LENGTH = 1000;

/**
 * Confidential message ("Vertrauliche Nachricht"): a text Send protected with the share password – the recipients
 * open the link from the mail and enter the password they get in the password mail or SMS. Vaultwarden deletes it
 * at `expiresAt` (1 to SEND_MAX_EXPIRE_DAYS days, chosen in the pane); no copy stays in Nextcloud or on this server.
 */
export async function createMessageSend(
  label: string,
  text: string,
  password: string,
  expiresAt: Date,
): Promise<{ url: string; id: string | undefined }> {
  const deletionDate = expiresAt.toISOString();
  const send = {
    name: label,
    notes: null,
    type: 0,
    text: { text, hidden: false },
    file: null,
    maxAccessCount: null,
    deletionDate,
    expirationDate: deletionDate,
    password,
    disabled: false,
    hideEmail: true,
  };
  const data = await withUnlock(() => bw<{ accessUrl?: string; id?: string }>("POST", "/object/send", send));
  if (!data?.accessUrl) throw new HttpError(502, t("server.vaultNoUrl"));
  return { url: data.accessUrl, id: typeof data.id === "string" ? data.id : undefined };
}
