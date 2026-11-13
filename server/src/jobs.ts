/**
 * Password jobs: one per Outlook draft; all files of that mail share password, Vaultwarden Send and password mail.
 *
 * The password mail is only sent once the main mail has really left Outlook:
 *   1. The add-in tags the draft with the job ID (Office.js custom property, travels into "Sent Items").
 *   2. On send, the add-in "arms" the job – the server stores a delegated Graph token, nothing is sent yet.
 *   3. A watcher looks for the tagged mail in the user's "Sent Items" and then sends the password mail to its
 *      actual recipients. If the mail never shows up (prompt cancelled, sending failed, draft discarded), the job
 *      is disarmed and stays open – further files are added to it, so there is only ever one password mail.
 *
 * Jobs are persisted (encrypted, DATA_DIR/password-jobs.json) so that a restart does not lose a pending mail.
 */
import { randomUUID } from "node:crypto";
import { readFileSync } from "node:fs";
import path from "node:path";
import type { FastifyBaseLogger } from "fastify";
import { config } from "./config.js";
import { findSentMessage, sendMail } from "./graph.js";
import { createAtomicWriter, seal, unseal, type Sealed } from "./storage.js";

export interface ArmedSend {
  /** Delegated Graph token (Mail.Read(.Shared) + Mail.Send(.Shared)) obtained when the user pressed "Send". */
  graphToken: string;
  tokenExpiresAt: number;
  armedAt: number;
  /** Shared mailbox in the "From" field (send as / on behalf); its "Sent Items" are searched as well. */
  mailbox?: string;
}

/** "Für jede Datei ein eigenes Kennwort": one Send and one password mail per file. */
export interface SeparatePart {
  fileName: string;
  sendUrl: string;
  sendId?: string;
  maxAccessCount: number;
  /** Its password mail went out (the others may still be pending after an error). */
  dispatched?: boolean;
}

export interface PasswordJob {
  owner: string;
  ownerUpn: string;
  fileNames: string[];
  /** Set for mails with one password per file; the top-level Send fields then describe the first file. */
  separate?: SeparatePart[];
  sendUrl: string;
  /** Vaultwarden Send ID – lets the number of accesses change until sending (jobs created before: none). */
  sendId?: string;
  /** Password mail only to this address instead of the mail's recipients ("Kennwort an andere Adresse"). */
  passwordTo?: string;
  /** How often the password link can be opened (jobs created before this field existed: unknown). */
  maxAccessCount?: number;
  dispatched: boolean;
  expiresAt: number;
  armed?: ArmedSend;
}

const MINUTE = 60_000;
const JOB_TTL = 7 * 24 * 60 * MINUTE;
/** How long after "Send" the watcher looks for the mail in "Sent Items". */
const ARM_WINDOW = 15 * MINUTE;
const POLL_INTERVAL = 20_000;
const AAD = "password-jobs";

const file = path.join(config.dataDir, "password-jobs.json");
const write = createAtomicWriter(file);
const jobs = new Map<string, PasswordJob>();

try {
  const sealed = JSON.parse(readFileSync(file, "utf8")) as Sealed;
  for (const [id, job] of Object.entries(JSON.parse(unseal(sealed, AAD)) as Record<string, PasswordJob>)) {
    jobs.set(id, job);
  }
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
    // Unreadable (e.g. CREDENTIAL_KEY changed) – start empty instead of refusing to start.
    console.warn(`password-jobs.json could not be read and is ignored: ${(err as Error).message}`);
  }
}

const persist = () => write(JSON.stringify(seal(JSON.stringify(Object.fromEntries(jobs)), AAD)));

export const getJob = (id: string) => jobs.get(id);

export async function createJob(
  owner: string,
  ownerUpn: string,
  fileName: string,
  sendUrl: string,
  maxAccessCount?: number,
  sendId?: string,
): Promise<string> {
  const id = randomUUID();
  jobs.set(id, {
    owner,
    ownerUpn,
    fileNames: [fileName],
    sendUrl,
    maxAccessCount,
    sendId,
    dispatched: false,
    expiresAt: Date.now() + JOB_TTL,
  });
  await persist();
  return id;
}

export async function addFileToJob(id: string, fileName: string): Promise<void> {
  jobs.get(id)!.fileNames.push(fileName);
  await persist();
}

/** First file of a mail with one password per file (see SeparatePart). */
export async function createSeparateJob(owner: string, ownerUpn: string, part: SeparatePart): Promise<string> {
  const id = randomUUID();
  jobs.set(id, {
    owner,
    ownerUpn,
    fileNames: [part.fileName],
    separate: [part],
    sendUrl: part.sendUrl,
    maxAccessCount: part.maxAccessCount,
    dispatched: false,
    expiresAt: Date.now() + JOB_TTL,
  });
  await persist();
  return id;
}

/** Further file with its own password and password mail. */
export async function addSeparatePart(id: string, part: SeparatePart): Promise<void> {
  const job = jobs.get(id)!;
  job.fileNames.push(part.fileName);
  (job.separate ??= []).push(part);
  await persist();
}

/** Alternative recipient of the password mail; undefined = the recipients of the mail. */
export async function setJobPasswordTo(id: string, passwordTo: string | undefined): Promise<void> {
  jobs.get(id)!.passwordTo = passwordTo;
  await persist();
}

/** New number of password-link accesses (the Send itself is updated by the caller). */
export async function setJobMaxAccess(id: string, maxAccessCount: number): Promise<void> {
  jobs.get(id)!.maxAccessCount = maxAccessCount;
  await persist();
}

/** Marks the job as "send pressed"; the watcher sends the password mail once the mail appears in "Sent Items". */
export async function armJob(id: string, armed: ArmedSend): Promise<void> {
  jobs.get(id)!.armed = armed;
  await persist();
}

const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

const accessText = (count?: number) =>
  count === undefined ? "begrenzte Anzahl Abrufe" : count === 1 ? "nur einmal abrufbar" : `bis zu ${count} Abrufe`;

/** The password mail goes to German-speaking recipients, so its text stays German. */
function passwordMailHtml(job: Pick<PasswordJob, "fileNames" | "sendUrl" | "maxAccessCount">, subject: string): string {
  const files = job.fileNames.map((f) => `<li>${escapeHtml(f)}</li>`).join("");
  return (
    `<p>Guten Tag,</p>` +
    `<p>für ${job.fileNames.length === 1 ? "die folgende Freigabe" : "die folgenden Freigaben"}` +
    (subject ? ` aus der Nachricht „${escapeHtml(subject)}“` : "") +
    ` benötigen Sie ein Kennwort:</p><ul>${files}</ul>` +
    `<p>Sie können es über den folgenden Link abrufen ` +
    `(${accessText(job.maxAccessCount)}, läuft nach ${config.sendExpireDays} Tagen ab):</p>` +
    `<p><a href="${escapeHtml(job.sendUrl)}">${escapeHtml(job.sendUrl)}</a></p>`
  );
}

/** One pass of the watcher over all jobs (exported for tests; normally driven by startJobWatcher). */
export async function checkJobs(log: Pick<FastifyBaseLogger, "info" | "warn">) {
  const now = Date.now();
  let changed = false;
  for (const [id, job] of jobs) {
    if (job.expiresAt < now) {
      jobs.delete(id);
      changed = true;
      continue;
    }
    const armed = job.armed;
    if (job.dispatched || !armed) continue;

    if (now - armed.armedAt > ARM_WINDOW || armed.tokenExpiresAt < now) {
      // The mail was not sent (prompt cancelled, sending failed, draft discarded) – keep the job open.
      job.armed = undefined;
      changed = true;
      log.info({ user: job.ownerUpn, jobId: id }, "Send not observed in Sent Items, job disarmed");
      continue;
    }

    try {
      // Small margin for clock differences between this server and Exchange.
      const since = armed.armedAt - 2 * MINUTE;
      // Exchange Online saves the copy in the sender's own "Sent Items" by default; depending on mailbox and client
      // settings it may (also or only) land in the shared mailbox's "Sent Items".
      let sent = await findSentMessage(armed.graphToken, id, since);
      if (!sent && armed.mailbox) {
        sent = await findSentMessage(armed.graphToken, id, since, armed.mailbox).catch((err) => {
          log.warn({ err, user: job.ownerUpn, mailbox: armed.mailbox }, "Shared mailbox not readable");
          return null;
        });
      }
      if (!sent) continue;
      // Same sender as the main mail – for a shared mailbox the password mail is sent as / on behalf of it.
      // "Kennwort an andere Adresse": only that address gets the password, not the recipients of the mail.
      const recipients = job.passwordTo ? { to: [job.passwordTo], cc: [], bcc: [] } : sent.recipients;
      const sender = sent.from ?? armed.mailbox;
      if (job.separate?.length) {
        // One password per file: one password mail per file, each naming its file. Parts already sent stay sent,
        // so a retry after an error does not send them twice.
        for (const part of job.separate) {
          if (part.dispatched) continue;
          await sendMail(armed.graphToken, recipients, `Kennwort für „${part.fileName}“: ${sent.subject || part.fileName}`,
            passwordMailHtml({ fileNames: [part.fileName], sendUrl: part.sendUrl, maxAccessCount: part.maxAccessCount },
              sent.subject), sender);
          part.dispatched = true;
          changed = true;
        }
      } else {
        await sendMail(armed.graphToken, recipients, `Kennwort für: ${sent.subject || job.fileNames[0]}`,
          passwordMailHtml(job, sent.subject), sender);
      }
      job.dispatched = true;
      job.armed = undefined;
      changed = true;
      log.info({ user: job.ownerUpn, from: sent.from, files: job.fileNames, other: Boolean(job.passwordTo) }, "Password mail sent");
    } catch (err) {
      log.warn({ err, user: job.ownerUpn, jobId: id }, "Password mail check failed, retrying");
    }
  }
  if (changed) await persist();
}

/** Starts the background check of armed jobs. */
export function startJobWatcher(log: FastifyBaseLogger) {
  let running = false;
  setInterval(async () => {
    if (running) return;
    running = true;
    try {
      await checkJobs(log);
    } catch (err) {
      log.error({ err }, "Job watcher failed");
    } finally {
      running = false;
    }
  }, POLL_INTERVAL).unref();
}
