/**
 * Download confirmation (DOWNLOAD_NOTICE=true): the sender gets a mail when a recipient downloads a file for the
 * first time. After the password went out (jobs.ts), every file link of the mail is registered here; a pass every
 * 10 minutes asks Nextcloud's Activity API with the owner's app password (the same data "Meine Freigaben" shows) and
 * counts only downloads after the registration. The mail is sent with CloudShare's own app token (application
 * permission Mail.Send), because the sender's delegated token has long expired by then.
 *
 * Upload notification (UPLOAD_NOTICE=true): upload links ("Anfordern") are registered the same way (`kind: "upload"`)
 * and stay registered until they expire; every pass lists their folder and mails the sender the files that arrived
 * since the last mail – at most one mail per link and pass, so several uploads in a row come as one mail.
 *
 * Privacy: no IP addresses or recipient data from Nextcloud are kept; entries are dropped when the link expired.
 */
import type { FastifyBaseLogger } from "fastify";
import { appGraphToken, HttpError } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import { locale, t } from "./i18n.js";
import { sendMailAsApp } from "./graph.js";
import { NextcloudClient, NOT_CONNECTED } from "./nextcloud.js";
import { PersistentMap } from "./db.js";
import { mailDocument, paragraph, TEXT_STYLE } from "./mailLayout.js";
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
  /** "upload": an upload link, watched for received files (missing: a file link, watched for its first download). */
  kind?: "upload";
  /** Upload link: the files already reported (or there before the link was sent). */
  knownFiles?: string[];
}

/** Is this kind of notice switched on (DOWNLOAD_NOTICE / UPLOAD_NOTICE)? */
const enabled = (n: Pick<Notice, "kind">) => (n.kind === "upload" ? config.uploadNotice : config.downloadNotice);
/** At most this many file names in an upload notification; the rest as "and N more". */
const MAX_NAMES = 10;

const INTERVAL = 10 * 60_000;
const DAY = 86_400_000;
/** Watched links in the database (db.ts, store "notices"). */
const notices = new PersistentMap<Notice>("notices");

/** Watches the given links (token → notice): file links for their first download, upload links for received files. */
export async function registerNotices(entries: Array<[string, Notice]>): Promise<void> {
  const wanted = entries.filter(([, notice]) => enabled(notice));
  if (wanted.length === 0) return;
  for (const [token, notice] of wanted) notices.set(token, notice);
}

/** Tests only. */
export const pendingNotices = () => [...notices.entries()];

const tokenOf = (url: string) => url.replace(/\/+$/, "").split("/").pop() ?? "";
const formatDay = (iso: string) => new Date(`${iso}T12:00:00Z`).toLocaleDateString(locale(), { timeZone: config.timeZone });
const formatTime = (iso: string) =>
  new Date(iso).toLocaleString(locale(), { timeZone: config.timeZone, dateStyle: "medium", timeStyle: "short" });

/** One pass (exported for tests; normally driven by startNotices). */
export async function runNotices(log: Pick<FastifyBaseLogger, "info" | "warn">, now = Date.now()) {
  if (notices.size === 0) return;
  for (const [token, n] of notices) {
    // Link expired (without a download), or the kind was switched off since it was registered.
    if (Date.parse(`${n.expireDate}T00:00:00Z`) + 2 * DAY < now || !enabled(n)) {
      notices.delete(token);
    }
  }
  const byOwner = new Map<string, Array<[string, Notice]>>();
  for (const item of notices) byOwner.set(item[1].owner, [...(byOwner.get(item[1].owner) ?? []), item]);

  let appToken: string | undefined;
  for (const [owner, items] of byOwner) {
    const creds = getCredentials(owner);
    if (!creds) {
      for (const [token] of items) notices.delete(token);
      continue;
    }
    const nc = new NextcloudClient(creds, () => removeCredentials(owner));
    try {
      const shares = new Map((await nc.listPublicShares()).map((s) => [tokenOf(s.url), s]));
      let uid: string | undefined;
      for (const [token, n] of items) {
        const share = shares.get(token);
        if (!share) continue; // revoked or already removed by Nextcloud – dropped once expired

        if (n.kind === "upload") {
          // Upload link: the files that arrived since the last mail; the link stays watched until it expires.
          uid ??= (await nc.userInfo()).id;
          const folder = share.path.split("/").filter(Boolean);
          const files = await nc.listFiles(uid, folder);
          const known = new Set(n.knownFiles ?? []);
          const received = (files ?? []).filter((name) => !known.has(name));
          if (received.length === 0) continue;
          appToken ??= await appGraphToken();
          const shown = received.slice(0, MAX_NAMES);
          const more = received.length - shown.length;
          const html = mailDocument({
            opening: paragraph(t("mail.uploadedText", { subject: n.subject || t("mail.noSubject"), recipients: n.recipients })),
            card: {
              accent: "success",
              title: t("mail.uploadedTitle", { count: received.length }),
              files: shown,
              button: { label: t("mail.uploadedButton"), url: `${config.nextcloudUrl}/index.php/apps/files/?dir=${encodeURIComponent(share.path)}` },
              details: [more > 0 ? t("mail.uploadedMore", { count: more }) : "", t("mail.uploadedDetails", { date: formatDay(n.expireDate) })]
                .filter(Boolean)
                .join(" · "),
            },
          });
          const subject = t("mail.uploadedSubject", { count: received.length, subject: n.subject || t("mail.noSubject") });
          await sendMailAsApp(appToken, owner, n.notifyTo, subject, html);
          n.knownFiles = [...known, ...received];
          notices.save(token);
          log.info({ user: n.ownerUpn, files: received.length }, "Upload notification sent");
          continue;
        }

        const downloads = await nc.publicDownloads(share.fileId);
        // Only downloads after the password went out (earlier links of the same file do not count).
        if (!downloads?.last || Date.parse(downloads.last) < n.registeredAt - 60_000) continue;
        appToken ??= await appGraphToken();
        const texts = templatesFor({ upn: n.ownerUpn, department: n.department });
        const values = {
          filename: n.fileName,
          subject: n.subject || t("mail.noSubject"),
          downloaded_at: formatTime(downloads.last),
          expires: formatDay(n.expireDate),
          recipients: n.recipients,
        };
        // Common layout (mailLayout.ts): the text of the template, then a green card with the file.
        const html = mailDocument({
          opening: renderHtml(texts.download_notice_body, values, TEXT_STYLE),
          card: {
            accent: "success",
            title: t("mail.downloadedTitle"),
            files: [n.fileName],
            details: t("mail.downloadedDetails", { time: values.downloaded_at, date: values.expires }),
          },
        });
        await sendMailAsApp(appToken, owner, n.notifyTo, render(texts.download_notice_subject, values), html);
        notices.delete(token);
        log.info({ user: n.ownerUpn, file: n.fileName }, "Download confirmation sent");
      }
    } catch (err) {
      if (err instanceof HttpError && err.code === NOT_CONNECTED) {
        for (const [token] of items) notices.delete(token);
      } else {
        log.warn({ err, user: items[0][1].ownerUpn }, "Download confirmation / upload notification check failed, retrying later");
      }
    }
  }
}

export function startNotices(log: FastifyBaseLogger) {
  if (!config.downloadNotice && !config.uploadNotice) return;
  let running = false;
  const pass = async () => {
    if (running) return;
    running = true;
    try {
      await runNotices(log);
    } catch (err) {
      log.error({ err }, "Download confirmation / upload notification pass failed");
    } finally {
      running = false;
    }
  };
  setInterval(pass, INTERVAL).unref();
}
