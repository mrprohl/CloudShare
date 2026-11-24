import { randomUUID } from "node:crypto";
import type { FastifyBaseLogger } from "fastify";
import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { PersistentMap } from "./db.js";
import { t } from "./i18n.js";

/*
 * The password of a step lives in a Vaultwarden Send (`bw serve`, service account): the password mail links to it, the
 * SMS and "Kennwort" in "Verlauf" read it from there. With VAULTWARDEN_ENABLED=false there is no Send: CloudShare keeps
 * the password in its own database instead (store "local-passwords", sealed with CREDENTIAL_KEY like every row) until
 * the links expire, and the password mail carries the password itself. The functions below take both: an ID starting
 * with "local:" is such an entry – also after Vaultwarden was switched on again, until it expires.
 */

/** A password kept by CloudShare itself (no Vaultwarden): until `keepUntil`, the expiry of its links. */
interface LocalPassword {
  password: string;
  keepUntil: number;
}
// Deleted when the links expire (purgeLocalPasswords), not by the general retention.
const localPasswords = new PersistentMap<LocalPassword>("local-passwords", { retention: false });
const LOCAL_PREFIX = "local:";
/** Whether a Send ID is a password kept by CloudShare itself. */
export const isLocalPassword = (id: string | undefined): boolean => Boolean(id?.startsWith(LOCAL_PREFIX));
/** Whether passwords go through Vaultwarden (VAULTWARDEN_ENABLED, default true). */
export const vaultwardenEnabled = (): boolean => config.vaultwardenEnabled;

/** Deletes the passwords whose links have expired (exported for tests). Returns how many. */
export function purgeLocalPasswords(now = Date.now()): number {
  let removed = 0;
  for (const [id, entry] of localPasswords) {
    if (entry.keepUntil < now) {
      localPasswords.delete(id);
      removed++;
    }
  }
  return removed;
}

/** Hourly pass over the passwords kept without Vaultwarden. */
export function startLocalPasswordPurge(log: Pick<FastifyBaseLogger, "info">): void {
  const pass = () => {
    const removed = purgeLocalPasswords();
    if (removed) log.info({ removed }, "Expired passwords deleted");
  };
  pass();
  setInterval(pass, 60 * 60_000).unref();
}


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

/** Waits at startup until `bw serve` is reachable and unlocks the vault (nothing to do without Vaultwarden). */
export async function initVault(log: (msg: string) => void, attempts = 30): Promise<void> {
  if (!vaultwardenEnabled()) {
    log("Vaultwarden off (VAULTWARDEN_ENABLED=false): passwords go in the password mail itself, kept sealed until the links expire");
    return;
  }
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
 * The password stored in a Send – for the SMS channel and for showing or sending it again ("Kennwort erneut senden").
 * The service account owns the Send, so it can read it until Vaultwarden deletes it (also after the recipients' access
 * ended). Without Vaultwarden: the password CloudShare keeps itself until the links expire.
 */
export async function readSendPassword(id: string): Promise<string> {
  if (isLocalPassword(id)) {
    const entry = localPasswords.get(id);
    if (!entry || entry.keepUntil < Date.now()) throw new HttpError(404, t("server.passwordUnavailable"));
    return entry.password;
  }
  const send = await withUnlock(() => bw<{ text?: { text?: string } }>("GET", `/object/send/${encodeURIComponent(id)}`));
  const password = send?.text?.text;
  if (!password) throw new HttpError(502, t("server.vaultNoPassword"));
  return password;
}

/**
 * How often a Send was opened ("Zustellnachweis": the password link or a confidential message) – null when Vaultwarden
 * no longer has it (deleted, expired) or there is no Send (the password went in the mail itself).
 */
export async function sendAccessCount(id: string): Promise<number | null> {
  if (isLocalPassword(id)) return null;
  try {
    const send = await withUnlock(() => bw<{ accessCount?: number }>("GET", `/object/send/${encodeURIComponent(id)}`));
    return typeof send?.accessCount === "number" ? send.accessCount : null;
  } catch {
    return null;
  }
}

/** Deletes a Send that is no longer needed (all links of its step were revoked) – or the password kept for it. */
export async function deleteSend(id: string): Promise<void> {
  if (isLocalPassword(id)) {
    localPasswords.delete(id);
    return;
  }
  await withUnlock(() => bw("DELETE", `/object/send/${encodeURIComponent(id)}`));
}

/**
 * Moves the end of the recipients' access to an existing Send to `expiresAt` – never past its deletion date, which
 * stays at the expiry of the links (password by SMS: the link is never sent, it only has to last until sending).
 * Nothing to move without a Send.
 */
export async function setSendExpiry(id: string, expiresAt: number): Promise<void> {
  if (isLocalPassword(id)) return;
  const path = `/object/send/${encodeURIComponent(id)}`;
  await withUnlock(async () => {
    const send = await bw<Record<string, unknown>>("GET", path);
    const deletion = Date.parse(String(send.deletionDate ?? ""));
    const end = Number.isFinite(deletion) ? Math.min(expiresAt, deletion) : expiresAt;
    await bw("PUT", path, { ...send, expirationDate: new Date(end).toISOString() });
  });
}

/** Changes how often an existing Send can be opened (until the mail is sent, "Kennwort abrufbar"); none without a Send. */
export async function setSendMaxAccess(id: string, maxAccessCount: number): Promise<void> {
  if (isLocalPassword(id)) return;
  const path = `/object/send/${encodeURIComponent(id)}`;
  await withUnlock(async () => {
    const send = await bw<Record<string, unknown>>("GET", path);
    await bw("PUT", path, { ...send, maxAccessCount });
  });
}

/**
 * Creates a hidden text Send containing the password and returns its URL, ID (allows changing the number of accesses
 * later and reading the password again) and the end of the recipients' access (ms).
 * Two dates: the recipients can open the link at most `maxAccessCount` times for `days` ("Kennwort-Link gültig für",
 * expiration); the Send itself lives until `keepUntil` – the expiry of the links – so the sender can still see and
 * resend the password (deletion; Vaultwarden allows at most 31 days ahead, so links with password last at most 30).
 * Without Vaultwarden no Send and no URL (`url` ""): the password is kept until `keepUntil` under a "local:" ID, and
 * the password mail carries it itself (passwordForMail).
 */
export async function createPasswordSend(
  label: string,
  password: string,
  maxAccessCount: number = config.sendMaxAccessCount,
  days: number = config.sendExpireDays,
  keepUntil?: number,
): Promise<{ url: string; id: string | undefined; expiresAt: number }> {
  const deletion = keepUntil ?? Date.now() + days * 86_400_000;
  if (!vaultwardenEnabled()) {
    const id = `${LOCAL_PREFIX}${randomUUID()}`;
    localPasswords.set(id, { password, keepUntil: deletion });
    return { url: "", id, expiresAt: deletion };
  }
  const expiresAt = Math.min(Date.now() + days * 86_400_000, deletion);
  const send = {
    name: label,
    notes: null,
    type: 0,
    text: { text: password, hidden: true },
    file: null,
    maxAccessCount,
    deletionDate: new Date(deletion).toISOString(),
    expirationDate: new Date(expiresAt).toISOString(),
    password: null,
    disabled: false,
    hideEmail: !config.sendShowCreator, // SEND_SHOW_CREATOR
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
 * Only with Vaultwarden – without it the pane does not offer "Nachricht".
 */
export async function createMessageSend(
  label: string,
  text: string,
  password: string,
  expiresAt: Date,
): Promise<{ url: string; id: string | undefined }> {
  if (!vaultwardenEnabled()) throw new HttpError(409, t("server.messagesUnavailable"), "NO_VAULTWARDEN");
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
    hideEmail: !config.sendShowCreator, // SEND_SHOW_CREATOR
  };
  const data = await withUnlock(() => bw<{ accessUrl?: string; id?: string }>("POST", "/object/send", send));
  if (!data?.accessUrl) throw new HttpError(502, t("server.vaultNoUrl"));
  return { url: data.accessUrl, id: typeof data.id === "string" ? data.id : undefined };
}
