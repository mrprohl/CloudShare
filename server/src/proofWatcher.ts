/**
 * "Zustellnachweis": watches the links of every open case (proof.ts) and records what happened since the last pass –
 * every download (Nextcloud activity), files received through an upload link, the password link or a confidential
 * message being opened (the access count of its Vaultwarden Send) and links that expired or disappeared. Every 10
 * minutes; a case is done once all its links ended, and is then only kept.
 */
import type { FastifyBaseLogger } from "fastify";
import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import { MESSAGE_PREFIX } from "./history.js";
import { NextcloudClient, NOT_CONNECTED } from "./nextcloud.js";
import { addEvent, caseSummary, endLink, openCases, saveCase, type ProofCase } from "./proof.js";
import { settingsFor } from "./settings.js";
import { tokenOf } from "./shares.js";
import { sendAccessCount } from "./vaultwarden.js";
import { sendDownloadProof, sendFinalProof, sendReminder } from "./proofMail.js";

const INTERVAL = 10 * 60_000;
const DAY = 86_400_000;
type Log = Pick<FastifyBaseLogger, "info" | "warn">;

const endOfDay = (date: string) => Date.parse(`${date}T00:00:00Z`) + DAY;

/** Access counts of the password Sends and messages: "password.fetched" / "message.opened" when they went up. */
async function watchSends(c: ProofCase, now: number) {
  const seen = (c.seenAccess ??= {});
  const messages = c.tokens.filter((t) => t.startsWith(MESSAGE_PREFIX) && !c.ended?.includes(t));
  const targets: Array<{ sendId: string; type: string; token?: string }> = [
    ...c.sendIds.map((sendId) => ({ sendId, type: "password.fetched" })),
    ...messages.map((token) => ({ sendId: token.slice(MESSAGE_PREFIX.length), type: "message.opened", token })),
  ];
  for (const { sendId, type, token } of targets) {
    const count = await sendAccessCount(sendId);
    if (count === null) {
      // A message whose Send is gone after its day: expired (revoking records it itself).
      if (token && endOfDay(c.links?.[token]?.expireDate ?? "1970-01-01") <= now) endLink(token, "expired", now);
      continue;
    }
    const before = seen[sendId] ?? 0;
    if (count <= before) continue;
    seen[sendId] = count;
    saveCase(c.id);
    addEvent(c.id, type, { count, previous: before, ...(token ? { links: [token] } : {}) }, now);
  }
}

/** Downloads and received files of the Nextcloud links of the owner's cases; links gone or expired are ended. */
async function watchLinks(owner: string, items: ProofCase[], now: number, log?: Log) {
  const creds = getCredentials(owner);
  if (!creds) return;
  const nc = new NextcloudClient(creds, () => removeCredentials(owner));
  const shares = new Map((await nc.listPublicShares()).map((s) => [tokenOf(s.url), s]));
  let uid: string | undefined;
  for (const c of items) {
    for (const [token, link] of Object.entries(c.links ?? {})) {
      if (link.kind === "message" || c.ended?.includes(token)) continue;
      const share = shares.get(token);
      if (!share) {
        endLink(token, endOfDay(link.expireDate) <= now ? "expired" : "gone", now);
        continue;
      }
      if (link.kind === "file") {
        const times = (await nc.publicDownloadTimes(share.fileId)) ?? [];
        const seen = new Set((c.seenDownloads ??= {})[token] ?? []);
        const fresh = times.filter((time) => time && !seen.has(time) && Date.parse(time) >= (link.createdAt ?? 0) - 60_000);
        if (fresh.length === 0) continue;
        c.seenDownloads[token] = [...seen, ...fresh];
        saveCase(c.id);
        for (const time of fresh) addEvent(c.id, "download", { links: [token], name: link.name, downloadedAt: time }, now);
        // The first download of the case: the sender gets the PDF (instead of the plain download confirmation).
        if (!c.downloadMailSent) {
          await sendDownloadProof(c, link.name, Date.parse(fresh[0]) || now)
            .then(() => {
              c.downloadMailSent = true;
              saveCase(c.id);
            })
            .catch((err) => log?.warn({ err, case: c.id }, "Delivery proof: PDF after the download not sent, retrying later"));
        }
      } else {
        uid ??= (await nc.userInfo()).id;
        const files = await nc.listFiles(uid, share.path.split("/").filter(Boolean));
        const seen = new Set((c.seenFiles ??= {})[token] ?? link.knownFiles ?? []);
        const fresh = (files ?? []).filter((name) => !seen.has(name));
        if (fresh.length === 0) continue;
        c.seenFiles[token] = [...seen, ...fresh];
        saveCase(c.id);
        addEvent(c.id, "files.received", { links: [token], files: fresh }, now);
      }
    }
  }
}

/**
 * Nothing fetched yet (no download, no files received, message not opened) and the last link ends within
 * PROOF_REMINDER_DAYS: the sender gets one reminder – unless they switched it off in the settings.
 */
async function remindIfNothingFetched(c: ProofCase, now: number, log?: Log) {
  if (c.reminderSent || config.proofReminderDays <= 0) return;
  const summary = caseSummary(c);
  if ((summary.state !== "sent" && summary.state !== "confirmed") || !summary.endsOn) return;
  if (endOfDay(summary.endsOn) - now > config.proofReminderDays * DAY) return;
  if (!settingsFor(c.owner).proofReminder) return;
  try {
    await sendReminder(c, summary.endsOn);
    c.reminderSent = true;
    saveCase(c.id);
  } catch (err) {
    log?.warn({ err, case: c.id }, "Delivery proof: reminder not sent, retrying later");
  }
}

/** Brings one case up to date right away (before its PDF is made) – downloads, received files, openings. */
export async function refreshCase(c: ProofCase, log: Log, now = Date.now()): Promise<void> {
  await watchLinks(c.owner, [c], now, log).catch((err) => log.warn({ err, case: c.id }, "Delivery proof: links not checked"));
  await watchSends(c, now).catch((err) => log.warn({ err, case: c.id }, "Delivery proof: Sends not checked"));
}

/** One pass (exported for tests; normally driven by startProofWatcher). */
export async function runProofWatch(log: Log, now = Date.now()) {
  const open = openCases().filter((c) => !c.done);
  const byOwner = new Map<string, ProofCase[]>();
  for (const c of open) byOwner.set(c.owner, [...(byOwner.get(c.owner) ?? []), c]);
  for (const [owner, items] of byOwner) {
    try {
      await watchLinks(owner, items, now, log);
    } catch (err) {
      if (!(err instanceof HttpError && err.code === NOT_CONNECTED)) {
        log.warn({ err, user: items[0].ownerUpn }, "Delivery proof: links not checked, retrying later");
      }
    }
    for (const c of items) {
      await watchSends(c, now).catch((err) => log.warn({ err, case: c.id }, "Delivery proof: Sends not checked"));
      await remindIfNothingFetched(c, now, log);
      // Done: every link ended – or two days past the last expiry (nothing can happen any more).
      const links = Object.entries(c.links ?? {});
      const lastExpiry = Math.max(0, ...links.map(([, l]) => endOfDay(l.expireDate)));
      if (links.length && (links.every(([t]) => c.ended?.includes(t)) || lastExpiry + 2 * DAY < now)) {
        // The complete PDF to the sender – before the case is done, so a failed mail is tried again next pass.
        if (!c.finalMailSent) {
          try {
            await sendFinalProof(c);
            c.finalMailSent = true;
          } catch (err) {
            log.warn({ err, case: c.id }, "Delivery proof: final PDF not sent, retrying later");
            continue;
          }
        }
        c.done = true;
        saveCase(c.id);
      }
    }
  }
}

export function startProofWatcher(log: FastifyBaseLogger) {
  if (!config.deliveryProof) return;
  let running = false;
  setInterval(async () => {
    if (running) return;
    running = true;
    try {
      await runProofWatch(log);
    } catch (err) {
      log.error({ err }, "Delivery proof pass failed");
    } finally {
      running = false;
    }
  }, INTERVAL).unref();
}
