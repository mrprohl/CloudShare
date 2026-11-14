/**
 * Deletes uploaded files after their link expired (CLEANUP_AFTER_DAYS, default 0 = right after expiry, "off" = never).
 *
 * Every upload has its own folder NEXTCLOUD_FOLDER/<date>_<id>. When a link to it is created, the folder and the
 * link's expiry date are registered here (encrypted in /data/cleanup.json, like the password jobs) – Nextcloud's own
 * housekeeping removes expired shares, so their expiry date cannot be read back later. An hourly pass deletes due
 * folders with the owner's app password; they go to the owner's Nextcloud trash. A folder that still has a valid
 * link (e.g. "Meine Freigaben → Einfügen" with a later date, or a link created in Nextcloud) is not deleted.
 * Received files (NEXTCLOUD_REQUEST_FOLDER) are never registered and therefore never deleted.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import type { FastifyBaseLogger } from "fastify";
import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import { NextcloudClient, NOT_CONNECTED } from "./nextcloud.js";
import { createAtomicWriter, seal, unseal, type Sealed } from "./storage.js";

export interface CleanupEntry {
  /** Entra oid of the owner (whose Nextcloud credentials delete the folder). */
  owner: string;
  /** [NEXTCLOUD_FOLDER, "<date>_<id>"] */
  folder: string[];
  /** Last day a link to this folder is valid (YYYY-MM-DD, the latest of all its links). */
  expireDate: string;
}

const DAY = 86_400_000;
const INTERVAL = 60 * 60_000;
const AAD = "cleanup";

const file = path.join(config.dataDir, "cleanup.json");
const write = createAtomicWriter(file);
const entries = new Map<string, CleanupEntry>();

try {
  const sealed = JSON.parse(readFileSync(file, "utf8")) as Sealed;
  for (const [key, entry] of Object.entries(JSON.parse(unseal(sealed, AAD)) as Record<string, CleanupEntry>)) {
    entries.set(key, entry);
  }
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
    console.warn(`cleanup.json could not be read and is ignored: ${(err as Error).message}`);
  }
}

const persist = () => write(JSON.stringify(seal(JSON.stringify(Object.fromEntries(entries)), AAD)));
const keyOf = (owner: string, folder: string[]) => `${owner}:${folder.join("/")}`;

/** Only upload folders directly below NEXTCLOUD_FOLDER – never anything else in the user's Nextcloud. */
export const isUploadFolder = (folder: string[]) =>
  folder.length === 2 && folder[0] === config.nextcloudFolder && /^\d{4}-\d{2}-\d{2}_[0-9a-f]{8}$/.test(folder[1]);

/** Registers (or extends) the expiry of a link to an uploaded file; `target` is the file's path segments. */
export async function registerForCleanup(owner: string, target: string[], expireDate: string): Promise<void> {
  const folder = target.slice(0, 2);
  if (!isUploadFolder(folder)) return;
  const key = keyOf(owner, folder);
  const existing = entries.get(key);
  if (existing && existing.expireDate >= expireDate) return;
  entries.set(key, { owner, folder, expireDate });
  await persist();
}

/**
 * A revoked link counts as expired today ("Zurückziehen"): with CLEANUP_AFTER_DAYS=0 the folder is deleted right away
 * (Nextcloud trash), otherwise it is due that many days from now. Returns what happened, for the taskpane's message.
 * `otherLinkActive`: another valid link still points into the folder – then nothing changes.
 */
export async function expireOnRevoke(
  owner: string,
  target: string[],
  otherLinkActive: boolean,
  deleteNow: (folder: string[]) => Promise<void>,
  now = Date.now(),
): Promise<"deleted" | "scheduled" | "kept"> {
  const folder = target.slice(0, 2);
  const afterDays = config.cleanupAfterDays;
  if (afterDays === null || otherLinkActive || !isUploadFolder(folder)) return "kept";
  const key = keyOf(owner, folder);
  if (afterDays === 0) {
    await deleteNow(folder);
    if (entries.delete(key)) await persist();
    return "deleted";
  }
  // Expired "yesterday", so that it becomes due in exactly `afterDays` days (see dueAt).
  const expireDate = new Date(now - DAY).toISOString().slice(0, 10);
  entries.set(key, { owner, folder, expireDate });
  await persist();
  return "scheduled";
}

/** Tests only. */
export const cleanupEntries = () => [...entries.values()];

/** Due: the link expired at the end of expireDate, plus the configured grace period. */
const dueAt = (entry: CleanupEntry, afterDays: number) => Date.parse(`${entry.expireDate}T00:00:00Z`) + (1 + afterDays) * DAY;

/** One pass (exported for tests; normally driven by startCleanup). */
export async function runCleanup(log: Pick<FastifyBaseLogger, "info" | "warn">, now = Date.now()) {
  const afterDays = config.cleanupAfterDays;
  if (afterDays === null) return;
  const due = [...entries.entries()].filter(([, e]) => dueAt(e, afterDays) <= now);
  const byOwner = new Map<string, Array<[string, CleanupEntry]>>();
  for (const item of due) byOwner.set(item[1].owner, [...(byOwner.get(item[1].owner) ?? []), item]);

  let changed = false;
  for (const [owner, items] of byOwner) {
    const creds = getCredentials(owner);
    if (!creds) {
      // Not connected any more – nothing can be deleted for this user; forget the entries.
      for (const [key] of items) entries.delete(key);
      changed = true;
      continue;
    }
    const nc = new NextcloudClient(creds, () => removeCredentials(owner));
    try {
      const { id: uid } = await nc.userInfo();
      const today = new Date(now).toISOString().slice(0, 10);
      const shares = await nc.listPublicShares();
      for (const [key, entry] of items) {
        const prefix = `/${entry.folder.join("/")}`;
        const active = shares.filter(
          (s) => (s.path === prefix || s.path.startsWith(`${prefix}/`)) && (s.expiration === null || s.expiration >= today),
        );
        if (active.length > 0) {
          // Still shared: postpone to the latest link, or keep for good if a link never expires.
          const dates = active.map((s) => s.expiration);
          if (dates.includes(null)) entries.delete(key);
          else entries.set(key, { ...entry, expireDate: (dates as string[]).sort().at(-1)! });
          changed = true;
          continue;
        }
        await nc.deletePath(uid, entry.folder);
        entries.delete(key);
        changed = true;
        log.info({ owner, folder: entry.folder.join("/"), expired: entry.expireDate }, "Expired upload deleted");
      }
    } catch (err) {
      if (err instanceof HttpError && err.code === NOT_CONNECTED) {
        for (const [key] of items) entries.delete(key);
        changed = true;
      } else {
        log.warn({ err, owner }, "Cleanup of expired uploads failed, retrying later");
      }
    }
  }
  if (changed) await persist();
}

/** Hourly pass; the first one a few minutes after start. */
export function startCleanup(log: FastifyBaseLogger) {
  if (config.cleanupAfterDays === null) {
    log.info("Cleanup of expired uploads is off (CLEANUP_AFTER_DAYS=off)");
    return;
  }
  let running = false;
  const pass = async () => {
    if (running) return;
    running = true;
    try {
      await runCleanup(log);
    } catch (err) {
      log.error({ err }, "Cleanup failed");
    } finally {
      running = false;
    }
  };
  setTimeout(pass, 5 * 60_000).unref();
  setInterval(pass, INTERVAL).unref();
}
