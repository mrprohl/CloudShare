/**
 * Download confirmation (DOWNLOAD_NOTICE=true): the sender gets a mail when a recipient downloads a file for the
 * first time. After the password went out (jobs.ts), every file link of the mail is registered here; a pass every
 * 10 minutes asks Nextcloud's Activity API with the owner's app password (the same data "Meine Freigaben" shows) and
 * counts only downloads after the registration. The mail is sent with CloudShare's own app token (application
 * permission Mail.Send), because the sender's delegated token has long expired by then.
 *
 * Privacy: no IP addresses or recipient data from Nextcloud are kept; entries are dropped when the link expired.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import type { FastifyBaseLogger } from "fastify";
import { appGraphToken, HttpError } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import { sendMailAsApp } from "./graph.js";
import { NextcloudClient, NOT_CONNECTED } from "./nextcloud.js";
import { createAtomicWriter, seal, unseal, type Sealed } from "./storage.js";
import { renderHtml, render, templatesFor } from "./templates.js";

export interface Notice {
  /** Entra oid (Nextcloud credentials) and mail address of the sender (recipient of the confirmation). */
  owner: string;
  ownerUpn: string;
  notifyTo: string;
  department?: string;
  fileName: string;
  subject: string;
  recipients: string;
  /** Last valid day of the link (YYYY-MM-DD). */
  expireDate: string;
  registeredAt: number;
}

const INTERVAL = 10 * 60_000;
const DAY = 86_400_000;
const AAD = "download-notices";
const file = path.join(config.dataDir, "download-notices.json");
const write = createAtomicWriter(file);
const notices = new Map<string, Notice>();

try {
  const sealed = JSON.parse(readFileSync(file, "utf8")) as Sealed;
  for (const [token, n] of Object.entries(JSON.parse(unseal(sealed, AAD)) as Record<string, Notice>)) notices.set(token, n);
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
    console.warn(`download-notices.json could not be read and is ignored: ${(err as Error).message}`);
  }
}

const persist = () => write(JSON.stringify(seal(JSON.stringify(Object.fromEntries(notices)), AAD)));

/** Watches the given file links (token → notice) for their first download. */
export async function registerNotices(entries: Array<[string, Notice]>): Promise<void> {
  if (!config.downloadNotice || entries.length === 0) return;
  for (const [token, notice] of entries) notices.set(token, notice);
  await persist();
}

/** Tests only. */
export const pendingNotices = () => [...notices.entries()];

const tokenOf = (url: string) => url.replace(/\/+$/, "").split("/").pop() ?? "";
const formatDay = (iso: string) => new Date(`${iso}T12:00:00Z`).toLocaleDateString("de-DE", { timeZone: "Europe/Berlin" });
const formatTime = (iso: string) =>
  new Date(iso).toLocaleString("de-DE", { timeZone: "Europe/Berlin", dateStyle: "medium", timeStyle: "short" });

/** One pass (exported for tests; normally driven by startNotices). */
export async function runNotices(log: Pick<FastifyBaseLogger, "info" | "warn">, now = Date.now()) {
  if (!config.downloadNotice || notices.size === 0) return;
  let changed = false;
  for (const [token, n] of notices) {
    if (Date.parse(`${n.expireDate}T00:00:00Z`) + 2 * DAY < now) {
      notices.delete(token); // link expired without a download
      changed = true;
    }
  }
  const byOwner = new Map<string, Array<[string, Notice]>>();
  for (const item of notices) byOwner.set(item[1].owner, [...(byOwner.get(item[1].owner) ?? []), item]);

  let appToken: string | undefined;
  for (const [owner, items] of byOwner) {
    const creds = getCredentials(owner);
    if (!creds) {
      for (const [token] of items) notices.delete(token);
      changed = true;
      continue;
    }
    const nc = new NextcloudClient(creds, () => removeCredentials(owner));
    try {
      const shares = new Map((await nc.listPublicShares()).map((s) => [tokenOf(s.url), s]));
      for (const [token, n] of items) {
        const share = shares.get(token);
        if (!share) continue; // revoked or already removed by Nextcloud – dropped once expired
        const downloads = await nc.publicDownloads(share.fileId);
        // Only downloads after the password went out (earlier links of the same file do not count).
        if (!downloads?.last || Date.parse(downloads.last) < n.registeredAt - 60_000) continue;
        appToken ??= await appGraphToken();
        const t = templatesFor({ upn: n.ownerUpn, department: n.department });
        const values = {
          filename: n.fileName,
          subject: n.subject || "(ohne Betreff)",
          downloaded_at: formatTime(downloads.last),
          expires: formatDay(n.expireDate),
          recipients: n.recipients,
        };
        const style = "margin:0 0 12px;font-family:'Segoe UI',-apple-system,Helvetica,Arial,sans-serif;font-size:14px;line-height:20px;color:#242424;";
        await sendMailAsApp(appToken, owner, n.notifyTo, render(t.download_notice_subject, values), renderHtml(t.download_notice_body, values, style));
        notices.delete(token);
        changed = true;
        log.info({ user: n.ownerUpn, file: n.fileName }, "Download confirmation sent");
      }
    } catch (err) {
      if (err instanceof HttpError && err.code === NOT_CONNECTED) {
        for (const [token] of items) notices.delete(token);
        changed = true;
      } else {
        log.warn({ err, user: items[0][1].ownerUpn }, "Download confirmation check failed, retrying later");
      }
    }
  }
  if (changed) await persist();
}

export function startNotices(log: FastifyBaseLogger) {
  if (!config.downloadNotice) return;
  let running = false;
  const pass = async () => {
    if (running) return;
    running = true;
    try {
      await runNotices(log);
    } catch (err) {
      log.error({ err }, "Download confirmation pass failed");
    } finally {
      running = false;
    }
  };
  setInterval(pass, INTERVAL).unref();
}
