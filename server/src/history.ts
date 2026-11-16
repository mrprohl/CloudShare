/**
 * What "Verlauf" needs beyond Nextcloud's own list of shares:
 *
 * - the mail each link went out with (subject, and the Graph ID of the sent mail so the pane can open it) – written
 *   when the watcher finds the sent mail (jobs.ts), per link token;
 * - the confidential messages ("Nachricht"): they are Vaultwarden Sends, not Nextcloud shares. Kept is only what lists
 *   them and lets the owner revoke them – the Send ID (to delete it), its access ID (the part of the link without the
 *   key, to find it in a mail being read) and its dates. Never the key or the text: CloudShare cannot read a message.
 *
 * Both live in the database (db.ts, stores "sent-mails" and "messages"); entries go when their link expired.
 */
import { PersistentMap } from "./db.js";

export interface SentMail {
  owner: string;
  /** Subject of the sent mail ("" if it had none). */
  subject: string;
  /** Graph ID (immutable) of the mail in the sender's own "Sent Items"; none for a shared mailbox. */
  mailId?: string;
  /** When it was sent (ms) and to whom (To and Cc – never Bcc): the tooltip of the subject in "Verlauf". */
  sentAt?: number;
  to?: string[];
  cc?: string[];
  /** When the link expires (ms) – the entry is dropped then. */
  keepUntil: number;
}

export interface MessageEntry {
  owner: string;
  /** Vaultwarden ID of the message's Send – to delete it ("Zurückziehen"). */
  sendId: string;
  /** Access ID in the link (…/#/send/<accessId>/<key>) – without the key. */
  accessId: string;
  createdAt: number;
  expiresAt: number;
}

const sentMails = new PersistentMap<SentMail>("sent-mails");
const messages = new PersistentMap<MessageEntry>("messages");

/** "Verlauf" token of a confidential message – shares its namespace with Nextcloud share tokens, never collides. */
export const MESSAGE_PREFIX = "msg-";
export const messageToken = (sendId: string) => `${MESSAGE_PREFIX}${sendId}`;
export const isMessageToken = (token: string) => token.startsWith(MESSAGE_PREFIX);

function prune() {
  const now = Date.now();
  for (const [token, mail] of sentMails) if (mail.keepUntil < now) sentMails.delete(token);
  for (const [token, message] of messages) if (message.expiresAt < now) messages.delete(token);
}

/**
 * The subject of the draft a link was inserted into – shown in "Verlauf" until the mail is sent (recordSentMail then
 * puts the real one and the mail's ID in its place). Never overwrites a known mail.
 */
export function recordDraftSubject(token: string, mail: Omit<SentMail, "mailId">): void {
  if (!sentMails.has(token)) sentMails.set(token, { ...mail });
}

/**
 * The draft's subject changed (the pane checks it every few seconds while it is open): the links of the mail not sent
 * yet show the new one – also when the draft is discarded afterwards. Links whose mail was found keep its subject.
 */
export function updateDraftSubject(links: Array<{ token: string; keepUntil: number }>, owner: string, subject: string): void {
  for (const { token, keepUntil } of links) {
    const mail = sentMails.get(token);
    if (mail && (mail.owner !== owner || mail.mailId)) continue;
    if (mail?.subject === subject) continue;
    sentMails.set(token, { owner, subject, keepUntil: mail?.keepUntil ?? keepUntil });
  }
}

/** Remembers the mail the links (tokens) went out with. */
export function recordSentMail(tokens: string[], mail: SentMail): void {
  prune();
  for (const token of tokens) sentMails.set(token, { ...mail });
}

export const sentMailFor = (token: string, owner: string): SentMail | undefined => {
  const mail = sentMails.get(token);
  return mail && mail.owner === owner ? mail : undefined;
};

/** Access ID of a Send link (https://vault…/#/send/<accessId>/<key>), or "" if it is not one. */
export const accessIdOf = (url: string) => /#\/send\/([^/]+)\//.exec(url)?.[1] ?? "";

/** Records a confidential message just created; returns its "Verlauf" token. */
export function recordMessage(entry: MessageEntry): string {
  prune();
  const token = messageToken(entry.sendId);
  messages.set(token, entry);
  return token;
}

/** The owner's messages that can still be read, newest first. */
export function messagesOf(owner: string): Array<MessageEntry & { token: string }> {
  const now = Date.now();
  return [...messages]
    .filter(([, m]) => m.owner === owner && m.expiresAt >= now)
    .map(([token, m]) => ({ ...m, token }))
    .sort((a, b) => b.createdAt - a.createdAt);
}

export const messageFor = (token: string, owner: string): MessageEntry | undefined => {
  const message = messages.get(token);
  return message && message.owner === owner && message.expiresAt >= Date.now() ? message : undefined;
};

export function forgetMessage(token: string): void {
  messages.delete(token);
  sentMails.delete(token);
}
