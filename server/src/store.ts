import type { LoginFlow } from "./nextcloud.js";

/**
 * Volatile storage for running uploads and pending logins (password jobs are persisted in jobs.ts).
 * A restart discards running uploads and logins; the user simply starts them again.
 */

export interface UploadState {
  owner: string;
  uid: string;
  target: string[];
  fileName: string;
  size: number;
  totalChunks: number;
  received: Set<number>;
  expiresAt: number;
}

export interface PendingConnect {
  owner: string;
  flow: LoginFlow;
  expiresAt: number;
}

export const uploads = new Map<string, UploadState>();
export const pendingConnects = new Map<string, PendingConnect>();

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
export const UPLOAD_TTL = 12 * HOUR;
/** Nextcloud discards login flow tokens after 20 minutes. */
export const CONNECT_TTL = 20 * MINUTE;

setInterval(() => {
  const now = Date.now();
  for (const [id, u] of uploads) if (u.expiresAt < now) uploads.delete(id);
  for (const [id, c] of pendingConnects) if (c.expiresAt < now) pendingConnects.delete(id);
}, 5 * MINUTE).unref();
