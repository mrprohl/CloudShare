/**
 * Where the password of each sent step lives – the ID of its Vaultwarden Send, never the password itself. Lets the
 * sender see the password again and send it once more ("Kennwort erneut senden": by SMS or mail) until the links
 * expire; then Vaultwarden deletes the Send and the entry is dropped. Keyed by the Send ID, looked up by the share
 * token (last segment of a link URL). Encrypted in /data/password-refs.json like the password jobs.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { config } from "./config.js";
import { createAtomicWriter, seal, unseal, type Sealed } from "./storage.js";

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
  /** Where the password mail went (recipients or the other address) – suggested for sending it again by mail. */
  mailTo: string[];
  /** Number the password went to by SMS (E.164) – suggested for sending it again by SMS; only the owner sees it. */
  sms?: string;
  /** Sender of the mail (shared mailbox) – a password mail sent again goes out the same way. */
  from?: string;
  /** Sends created when the password was sent again by mail – deleted together with the first one. */
  resentSends?: string[];
  maxAccessCount: number;
  senderName?: string;
  department?: string;
}

const AAD = "password-refs";
const file = path.join(config.dataDir, "password-refs.json");
const write = createAtomicWriter(file);
const refs = new Map<string, PasswordRef>();

try {
  const sealed = JSON.parse(readFileSync(file, "utf8")) as Sealed;
  for (const [id, ref] of Object.entries(JSON.parse(unseal(sealed, AAD)) as Record<string, PasswordRef>)) {
    refs.set(id, ref);
  }
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
    console.warn(`password-refs.json could not be read and is ignored: ${(err as Error).message}`);
  }
}

function persist() {
  const now = Date.now();
  for (const [id, ref] of refs) if (ref.keepUntil < now || ref.tokens.length === 0) refs.delete(id);
  return write(JSON.stringify(seal(JSON.stringify(Object.fromEntries(refs)), AAD)));
}

/** Records the Send of a step once its mail was sent (again for the same Send: the newer data wins). */
export async function registerPasswordRef(ref: PasswordRef): Promise<void> {
  refs.set(ref.sendId, ref);
  await persist();
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
  change: { mailTo?: string[]; sms?: string; resentSend?: string },
): Promise<void> {
  const ref = refs.get(sendId);
  if (!ref) return;
  if (change.mailTo) ref.mailTo = change.mailTo;
  if (change.sms) ref.sms = change.sms;
  if (change.resentSend) (ref.resentSends ??= []).push(change.resentSend);
  await persist();
}

/**
 * A link was revoked: it no longer needs the password. Returns the IDs of the step's Sends when no link of it is
 * left – the caller deletes them then.
 */
export async function forgetPasswordToken(token: string, owner: string): Promise<string[]> {
  const ref = passwordRefFor(token, owner);
  if (!ref) return [];
  ref.tokens = ref.tokens.filter((t) => t !== token);
  await persist();
  return ref.tokens.length === 0 ? [ref.sendId, ...(ref.resentSends ?? [])] : [];
}

/** Entries holding these tokens (seven.io reported an SMS of theirs as not delivered). */
export const passwordRefsForTokens = (tokens: string[]) =>
  [...refs.values()].filter((ref) => ref.tokens.some((t) => tokens.includes(t)));
