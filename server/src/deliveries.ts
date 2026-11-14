/**
 * How the password of each link went out – shown in "Meine Freigaben" ("Kennwort per SMS an +49 170 •••• 78 –
 * zugestellt") and updated by seven.io's delivery reports. Keyed by the share token (last segment of the link URL).
 * Encrypted in /data/deliveries.json like the password jobs; entries are dropped after DELIVERY_TTL.
 */
import { readFileSync } from "node:fs";
import path from "node:path";
import { config } from "./config.js";
import type { SmsStatus } from "./sms.js";
import { createAtomicWriter, seal, unseal, type Sealed } from "./storage.js";

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
const AAD = "deliveries";
const file = path.join(config.dataDir, "deliveries.json");
const write = createAtomicWriter(file);
const deliveries = new Map<string, Delivery>();

try {
  const sealed = JSON.parse(readFileSync(file, "utf8")) as Sealed;
  for (const [token, d] of Object.entries(JSON.parse(unseal(sealed, AAD)) as Record<string, Delivery>)) {
    deliveries.set(token, d);
  }
} catch (err) {
  if ((err as NodeJS.ErrnoException).code !== "ENOENT") {
    console.warn(`deliveries.json could not be read and is ignored: ${(err as Error).message}`);
  }
}

function persist() {
  const now = Date.now();
  for (const [token, d] of deliveries) if (now - d.at > DELIVERY_TTL) deliveries.delete(token);
  return write(JSON.stringify(seal(JSON.stringify(Object.fromEntries(deliveries)), AAD)));
}

/** Records the delivery for every token of a step (one password message covers all files of a step). */
export async function recordDelivery(tokens: string[], delivery: Delivery): Promise<void> {
  for (const token of tokens) deliveries.set(token, { ...delivery });
  await persist();
}

export const deliveryFor = (token: string, owner: string) => {
  const d = deliveries.get(token);
  return d && d.owner === owner ? d : undefined;
};

/** Delivery report of seven.io (idempotent: the same report twice changes nothing). Returns false if unknown. */
export async function updateSmsStatus(messageId: string, status: SmsStatus, at = Date.now()): Promise<boolean> {
  let found = false;
  for (const d of deliveries.values()) {
    if (d.messageId !== messageId) continue;
    found = true;
    // A final state is not overwritten by a late intermediate report.
    if (d.status === "delivered" || d.status === "failed") continue;
    d.status = status;
    if (status === "delivered") d.deliveredAt = at;
  }
  if (found) await persist();
  return found;
}
