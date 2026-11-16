/**
 * How the password of each link went out – shown in "Meine Freigaben" ("Kennwort per SMS an +49 170 •••• 78 –
 * zugestellt") and updated by seven.io's delivery reports. Keyed by the share token (last segment of the link URL).
 * Stored in the database (db.ts, store "deliveries"); entries are dropped after DELIVERY_TTL.
 */
import { PersistentMap } from "./db.js";
import type { SmsStatus } from "./sms.js";

export interface Delivery {
  owner: string;
  channel: "mail" | "sms";
  /** Recipient for display: mail addresses or the masked phone number. Never the full number. */
  to: string;
  at: number;
  /** SMS only. */
  messageId?: string;
  status?: SmsStatus;
  deliveredAt?: number;
}

const DELIVERY_TTL = 120 * 86_400_000;
const deliveries = new PersistentMap<Delivery>("deliveries");

/** Drops entries older than DELIVERY_TTL. */
function prune() {
  const now = Date.now();
  for (const [token, d] of deliveries) if (now - d.at > DELIVERY_TTL) deliveries.delete(token);
}

/** Records the delivery for every token of a step (one password message covers all files of a step). */
export async function recordDelivery(tokens: string[], delivery: Delivery): Promise<void> {
  prune();
  for (const token of tokens) deliveries.set(token, { ...delivery });
}

export const deliveryFor = (token: string, owner: string) => {
  const d = deliveries.get(token);
  return d && d.owner === owner ? d : undefined;
};

/**
 * Delivery report of seven.io (idempotent: the same report twice changes nothing). `known`: the message ID belongs to
 * a recorded SMS; `failedTokens`: links whose SMS just turned out as not delivered (the sender is told once).
 */
export async function updateSmsStatus(
  messageId: string,
  status: SmsStatus,
  at = Date.now(),
): Promise<{ known: boolean; failedTokens: string[] }> {
  let known = false;
  const failedTokens: string[] = [];
  for (const [token, d] of deliveries) {
    if (d.messageId !== messageId) continue;
    known = true;
    // A final state is not overwritten by a late intermediate report.
    if (d.status === "delivered" || d.status === "failed") continue;
    d.status = status;
    if (status === "delivered") d.deliveredAt = at;
    if (status === "failed") failedTokens.push(token);
    deliveries.save(token);
  }
  return { known, failedTokens };
}
