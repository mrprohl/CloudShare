/**
 * "Zustellnachweis" (proof of delivery) – only for mails whose sender ticked it in "Gültigkeit und Sprache".
 *
 * One case per mail, named by a tracking ID ("20261117-1146-7Q2F") that the link card, the password mail, the
 * confirmations and "Verlauf" show. The case records what happened, each event with its time:
 *
 *   case.opened       who, from which client (Outlook host, platform, version)
 *   link.created      what went out (file name, size, SHA-256 of the content), protection, expiry, password ways
 *   message.sent         subject, Message-ID, sender, To/Cc (number of Bcc), time – from "Sent Items"
 *   password.sent     way and recipient (SMS masked), also when sent again
 *   sms.status        delivery report of seven.io
 *   password.fetched  the password link was opened (Vaultwarden's access count went up)
 *   download          a download through the link (Nextcloud activity, every one)
 *   files.received    files arrived through an upload link
 *   revoked / expired the link ended
 *
 * Every event carries the SHA-256 of the previous one (a hash chain): changing or removing an event afterwards
 * breaks the chain, which verifyCase() and the PDF show. Cases are kept PROOF_RETENTION_MONTHS after their last event
 * (default 3 years), independent of DATA_RETENTION_DAYS. Store "proofs" in the database (sealed like everything).
 */
import { createHash, randomInt } from "node:crypto";
import { config } from "./config.js";
import { PersistentMap } from "./db.js";

export interface ClientInfo {
  /** Office host and platform ("Outlook", "PC" / "OfficeOnline" / "Mac"), Outlook's version, the add-in language. */
  host?: string;
  platform?: string;
  version?: string;
  language?: string;
}

export interface ProofEvent {
  at: number;
  type: string;
  data: Record<string, unknown>;
  /** Hash of the previous event ("" for the first) and of this one: sha256(prev + canonical JSON of at, type, data). */
  prev: string;
  hash: string;
}

export interface ProofLink {
  /** Where the gateway (/r/s/<token>) sends a verified recipient: the Nextcloud link. Messages have none. */
  url?: string;
  name: string;
  kind: "file" | "request" | "message";
  /** Last valid day (YYYY-MM-DD). */
  expireDate: string;
  /** SHA-256 of the file (hex), when it was uploaded through CloudShare. */
  sha256?: string;
  /** When the link was created (ms) – downloads of the same file through earlier links do not count. */
  createdAt?: number;
  /** Upload link: files already in the folder when it was created – not reported as received. */
  knownFiles?: string[];
}

export interface ProofCase {
  id: string;
  owner: string;
  ownerUpn: string;
  createdAt: number;
  /** Language of the recipients (gateway page, code mail) – the one chosen for the first link. */
  language?: string;
  /** Language of the sender's add-in when the case was opened (the PDF's default). */
  senderLanguage?: string;
  /** The links of the case by token. */
  links?: Record<string, ProofLink>;
  /** Addresses the mail went to (To, Cc, Bcc; lower case) – only they can get a code at the gateway. Known once sent. */
  recipients?: string[];
  /** Internet Message-IDs already recorded as sent (a retry of the job does not record the mail twice). */
  mailIds?: string[];
  /** Share tokens (Nextcloud links, "msg-…" messages) of the case. */
  tokens: string[];
  /** Vaultwarden Sends with the password of the case's steps – their access count is watched. */
  sendIds: string[];
  events: ProofEvent[];
  /** Watcher state: access count per Send, download times per token already recorded, ended links. */
  seenAccess?: Record<string, number>;
  seenDownloads?: Record<string, string[]>;
  seenFiles?: Record<string, string[]>;
  ended?: string[];
  /** Nothing left to watch (all links ended or expired) – the case is only kept. */
  done?: boolean;
  /** The sender got the PDF after the first download / at the end (proofMail.ts). */
  downloadMailSent?: boolean;
  finalMailSent?: boolean;
  /** Held-back passwords already released per Send ID (proofPassword.ts): addresses mailed, other addresses, SMS. */
  released?: Record<string, { mailed: string[]; others?: boolean; sms?: boolean }>;
}

const cases = new PersistentMap<ProofCase>("proofs", { retention: false });

/** The day a case may go: its last event plus PROOF_RETENTION_MONTHS calendar months. */
const keptUntil = (last: number) => {
  const date = new Date(last);
  date.setUTCMonth(date.getUTCMonth() + config.proofRetentionMonths);
  return date.getTime();
};

/** "20261117-1146-7Q2F": date and time (server time zone) and four characters without look-alikes (no 0/O, 1/I). */
export function newTrackingId(now = new Date()): string {
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: config.timeZone, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false,
  }).formatToParts(now);
  const get = (type: string) => parts.find((p) => p.type === type)?.value ?? "00";
  const alphabet = "23456789ABCDEFGHJKLMNPQRSTUVWXYZ";
  const random = Array.from({ length: 4 }, () => alphabet[randomInt(alphabet.length)]).join("");
  const id = `${get("year")}${get("month")}${get("day")}-${get("hour").replace("24", "00")}${get("minute")}-${random}`;
  return cases.has(id) ? newTrackingId(now) : id;
}

/** Canonical JSON: keys sorted, so the hash does not depend on the order fields were written in. */
function canonical(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(canonical).join(",")}]`;
  if (value && typeof value === "object") {
    return `{${Object.keys(value as Record<string, unknown>)
      .sort()
      .filter((k) => (value as Record<string, unknown>)[k] !== undefined)
      .map((k) => `${JSON.stringify(k)}:${canonical((value as Record<string, unknown>)[k])}`)
      .join(",")}}`;
  }
  return JSON.stringify(value);
}

const eventHash = (prev: string, e: Pick<ProofEvent, "at" | "type" | "data">) =>
  createHash("sha256").update(prev).update(canonical({ at: e.at, type: e.type, data: e.data })).digest("hex");

function prune() {
  const now = Date.now();
  for (const [id, c] of cases) {
    const last = c.events.at(-1)?.at ?? c.createdAt;
    if (keptUntil(last) < now) cases.delete(id);
  }
}

/** Adds an event to a case (the chain goes on from its last one). */
export function addEvent(id: string, type: string, data: Record<string, unknown>, at = Date.now()): void {
  const c = cases.get(id);
  if (!c) return;
  const prev = c.events.at(-1)?.hash ?? "";
  const event = { at, type, data };
  c.events.push({ ...event, prev, hash: eventHash(prev, event) });
  cases.save(id);
}

/** Only short, plain values of the client description (Office.js diagnostics) are kept. */
export function clientInfo(raw: unknown): ClientInfo {
  const value = (raw ?? {}) as Record<string, unknown>;
  const text = (v: unknown) => (typeof v === "string" && /^[\w .()/+-]{1,40}$/.test(v) ? v : undefined);
  return { host: text(value.host), platform: text(value.platform), version: text(value.version), language: text(value.language) };
}

/** Opens a case for a mail (the first link with "Zustellnachweis"); returns its tracking ID. */
export function openCase(
  owner: string,
  ownerUpn: string,
  client: ClientInfo | undefined,
  languages: { recipients?: string; sender?: string } = {},
): string {
  prune();
  const id = newTrackingId();
  cases.set(id, {
    id, owner, ownerUpn, createdAt: Date.now(), tokens: [], sendIds: [], events: [], links: {},
    language: languages.recipients, senderLanguage: languages.sender,
  });
  addEvent(id, "case.opened", { sender: ownerUpn, client: client ?? {} });
  return id;
}

/**
 * The case a new link joins: the mail's case (its tracking ID from the draft) or a new one. Undefined when the
 * sender did not tick "Zustellnachweis".
 */
export function caseForLink(
  wanted: unknown,
  requestedId: unknown,
  owner: string,
  ownerUpn: string,
  client: unknown,
  languages: { recipients?: string; sender?: string },
): string | undefined {
  if (wanted !== true) return undefined;
  const existing = caseFor(requestedId, owner);
  // A case whose mail went out already belongs to that mail – a new mail gets its own.
  if (existing && !existing.mailIds?.length) return existing.id;
  return openCase(owner, ownerUpn, clientInfo(client), languages);
}

/** The case of `owner` with this ID (undefined: none, someone else's or gone). */
export const caseFor = (id: unknown, owner: string): ProofCase | undefined => {
  const c = typeof id === "string" ? cases.get(id) : undefined;
  return c && c.owner === owner ? c : undefined;
};

/** The case a link (share token) belongs to. */
export const caseOfToken = (token: string): ProofCase | undefined => {
  for (const c of cases.values()) if (c.tokens.includes(token)) return c;
  return undefined;
};

/** Records a link of the case: what went out and how it is protected. */
export function addLink(
  id: string,
  token: string,
  link: ProofLink,
  sendId: string | undefined,
  data: Record<string, unknown>,
): void {
  const c = cases.get(id);
  if (!c) return;
  if (!c.tokens.includes(token)) c.tokens.push(token);
  (c.links ??= {})[token] = link;
  if (sendId && !c.sendIds.includes(sendId)) c.sendIds.push(sendId);
  cases.save(id);
  addEvent(id, "link.created", { token, kind: link.kind, name: link.name, expireDate: link.expireDate, ...data });
}

/** The address of the gateway for a link of a case: the recipient confirms their address there first. */
export const gatewayUrl = (token: string) => `${config.publicBaseUrl}/r/s/${encodeURIComponent(token)}`;

/** The link card and "Verlauf" show the gateway instead of the Nextcloud link for links of a case. */
export function linkUrlFor(token: string, nextcloudUrl: string): string {
  const c = caseOfToken(token);
  return c && c.links?.[token]?.url ? gatewayUrl(token) : nextcloudUrl;
}

/**
 * The mail with the links was found in "Sent Items" (jobs.ts): its Message-ID, sender and recipients go into every case
 * the links belong to – once per mail – and the recipients may now confirm their address at the gateway.
 */
export function recordMailSent(
  tokens: string[],
  mail: { subject: string; messageId?: string; from: string | null; to: string[]; cc: string[]; bcc: string[]; sentAt?: string },
): void {
  const ids = new Set(tokens.map((t) => caseOfToken(t)?.id).filter((id): id is string => Boolean(id)));
  for (const id of ids) {
    const c = cases.get(id)!;
    const key = mail.messageId ?? mail.sentAt ?? "sent";
    if (c.mailIds?.includes(key)) continue;
    (c.mailIds ??= []).push(key);
    const all = [...mail.to, ...mail.cc, ...mail.bcc].map((a) => a.toLowerCase());
    c.recipients = [...new Set([...(c.recipients ?? []), ...all])];
    cases.save(id);
    addEvent(id, "message.sent", {
      subject: mail.subject,
      messageId: mail.messageId ?? null,
      from: mail.from,
      to: mail.to,
      cc: mail.cc,
      bcc: mail.bcc,
      sentAt: mail.sentAt ?? null,
      links: tokens.filter((t) => c.tokens.includes(t)),
    });
  }
}

/** An event for every case the links (tokens) belong to – once per case. */
export function eventForTokens(tokens: string[], type: string, data: Record<string, unknown>, at = Date.now()): void {
  const ids = new Set(tokens.map((t) => caseOfToken(t)?.id).filter((id): id is string => Boolean(id)));
  for (const id of ids) {
    const c = cases.get(id)!;
    addEvent(id, type, { ...data, links: tokens.filter((t) => c.tokens.includes(t)) }, at);
  }
}

/** Whether the chain of a case is intact: every event hashes to its stored hash and points to the one before. */
export function verifyCase(c: ProofCase): { intact: boolean; brokenAt?: number } {
  let prev = "";
  for (let i = 0; i < c.events.length; i++) {
    const e = c.events[i];
    if (e.prev !== prev || eventHash(prev, e) !== e.hash) return { intact: false, brokenAt: i };
    prev = e.hash;
  }
  return { intact: true };
}

/** Open cases for the watcher (proofWatcher.ts): every case not yet past its retention. */
export const openCases = () => {
  prune();
  return [...cases.values()];
};

/** Saves watcher state changed in place. */
export const saveCase = (id: string) => cases.save(id);

/** Tests only. */
export const allCases = () => [...cases.values()];

/** The case and link of a token for the notes to the sender: tracking ID and the file's fingerprint. */
export function proofOf(token: string): { id: string; sha256?: string } | undefined {
  const c = caseOfToken(token);
  return c && { id: c.id, sha256: c.links?.[token]?.sha256 };
}

/** A link of a case ended: "revoked" by the sender (or for a mail never sent), "expired", or "deleted" by Nextcloud. */
export function endLink(token: string, reason: "revoked" | "expired" | "gone", at = Date.now()): void {
  const c = caseOfToken(token);
  if (!c || c.ended?.includes(token)) return;
  (c.ended ??= []).push(token);
  cases.save(c.id);
  addEvent(c.id, `link.${reason}`, { links: [token] }, at);
}

/** A case by its tracking ID (the signed PDF address was checked before). */
export const caseById = (id: string): ProofCase | undefined => cases.get(id);

/** All cases of `owner`, newest first – also the ones whose links ended long ago (kept PROOF_RETENTION_MONTHS). */
export const casesOf = (owner: string): ProofCase[] => {
  prune();
  return [...cases.values()].filter((c) => c.owner === owner).sort((a, b) => b.createdAt - a.createdAt);
};
