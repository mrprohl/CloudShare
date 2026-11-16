/**
 * Where the password of each sent step lives – the ID of its Vaultwarden Send, never the password itself. Lets the
 * sender see the password again and send it once more ("Kennwort erneut senden": by SMS or mail) until the links
 * expire; then Vaultwarden deletes the Send and the entry is dropped. Keyed by the Send ID, looked up by the share
 * token (last segment of a link URL). Stored in the database (db.ts, store "password-refs").
 */
import { PersistentMap } from "./db.js";
import type { MailRecipients } from "./graph.js";

export interface PasswordRef {
  owner: string;
  ownerUpn: string;
  sendId: string;
  /** Share tokens of the step's links. */
  tokens: string[];
  fileNames: string[];
  /** Expiry of the links (ms): Vaultwarden deletes the Send then, the entry is dropped. */
  keepUntil: number;
  /** Subject of the sent mail ("" if it had none). */
  subject: string;
  /**
   * Recipients of the sent mail as they were (To, Cc, Bcc) – "Empfänger der Nachricht" sends the password again with
   * exactly this distribution. Entries written before it existed have only `mailTo`.
   */
  recipients?: MailRecipients;
  /** Address(es) the password went to instead of the recipients ("Andere E-Mail-Adresse") – suggested again. */
  otherTo?: string[];
  /** Entries before `recipients`: To and Cc of the mail (or the other address) – fallback only. */
  mailTo?: string[];
  /** Number the password went to by SMS (E.164) – suggested for sending it again by SMS; only the owner sees it. */
  sms?: string;
  /** Sender of the mail (shared mailbox) – a password mail sent again goes out the same way. */
  from?: string;
  /** Sends created when the password was sent again by mail – deleted together with the first one. */
  resentSends?: string[];
  maxAccessCount: number;
  senderName?: string;
  department?: string;
  /** Language the recipients got the password in – sending it again uses it too. */
  language?: string;
  /** Language of the sender's add-in when the mail was sent – a later note about an undelivered SMS comes in it. */
  senderLanguage?: string;
}

const refs = new PersistentMap<PasswordRef>("password-refs");

/** Drops entries whose links expired or were all revoked. */
function prune() {
  const now = Date.now();
  for (const [id, ref] of refs) if (ref.keepUntil < now || ref.tokens.length === 0) refs.delete(id);
}

/** Records the Send of a step once its mail was sent (again for the same Send: the newer data wins). */
export async function registerPasswordRef(ref: PasswordRef): Promise<void> {
  prune();
  refs.set(ref.sendId, ref);
}

/** The entry of a link of `owner`, if its password can still be read (null: none, expired or someone else's). */
export function passwordRefFor(token: string, owner: string): PasswordRef | null {
  const now = Date.now();
  for (const ref of refs.values()) {
    if (ref.owner === owner && ref.keepUntil >= now && ref.tokens.includes(token)) return ref;
  }
  return null;
}

/** Remembers where the password went last and the Send of a password mail sent again. */
export async function updatePasswordRef(
  sendId: string,
  change: { otherTo?: string[]; sms?: string; resentSend?: string },
): Promise<void> {
  const ref = refs.get(sendId);
  if (!ref) return;
  if (change.otherTo) ref.otherTo = change.otherTo;
  if (change.sms) ref.sms = change.sms;
  if (change.resentSend) (ref.resentSends ??= []).push(change.resentSend);
  refs.save(sendId);
}

/**
 * A link was revoked: it no longer needs the password. Returns the IDs of the step's Sends when no link of it is
 * left – the caller deletes them then.
 */
export async function forgetPasswordToken(token: string, owner: string): Promise<string[]> {
  const ref = passwordRefFor(token, owner);
  if (!ref) return [];
  ref.tokens = ref.tokens.filter((t) => t !== token);
  if (ref.tokens.length === 0) refs.delete(ref.sendId);
  else refs.save(ref.sendId);
  return ref.tokens.length === 0 ? [ref.sendId, ...(ref.resentSends ?? [])] : [];
}

/** Entries holding these tokens (seven.io reported an SMS of theirs as not delivered). */
export const passwordRefsForTokens = (tokens: string[]) =>
  [...refs.values()].filter((ref) => ref.tokens.some((t) => tokens.includes(t)));

/** The original distribution of the mail (To, Cc, Bcc); older entries: their `mailTo` as To. */
export const originalRecipients = (ref: PasswordRef): MailRecipients =>
  ref.recipients ?? { to: ref.mailTo ?? [], cc: [], bcc: [] };
