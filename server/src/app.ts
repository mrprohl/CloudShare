import rateLimit from "@fastify/rate-limit";
import fastifyStatic from "@fastify/static";
import Fastify, { type FastifyRequest } from "fastify";
import { decodeJwt } from "jose";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { authenticate, HttpError, onBehalfOf, type User } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials, saveCredentials } from "./credentials.js";
import { fetchProfile, GRAPH_PROFILE_SCOPES, GRAPH_SCOPES } from "./graph.js";
import { addFileToJob, armJob, createJob, getJob } from "./jobs.js";
import {
  NextcloudClient,
  NOT_CONNECTED,
  pollLoginFlow,
  QUOTA_EXCEEDED,
  startLoginFlow,
  type PublicShare,
} from "./nextcloud.js";
import { clampInt } from "./limits.js";
import { passwordProblem } from "./password.js";
import { CONNECT_TTL, pendingConnects, UPLOAD_TTL, uploads, type UploadState } from "./store.js";
import { createPasswordSend } from "./vaultwarden.js";

function nextcloudFor(user: User): NextcloudClient {
  const creds = getCredentials(user.oid);
  if (!creds) throw new HttpError(409, "Nextcloud ist noch nicht verbunden.", NOT_CONNECTED);
  return new NextcloudClient(creds, () => removeCredentials(user.oid));
}

function sanitizeFileName(name: unknown): string {
  const cleaned = String(name ?? "")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .trim()
    .slice(0, 200);
  if (!cleaned || cleaned === "." || cleaned === "..") throw new HttpError(400, "Ungültiger Dateiname");
  return cleaned;
}

function ownUpload(id: string, user: User): UploadState {
  const upload = uploads.get(id);
  if (!upload || upload.owner !== user.oid) throw new HttpError(404, "Upload nicht gefunden");
  return upload;
}

interface ShareRequest {
  password?: unknown;
  expireDays?: unknown;
  /** How often the password link (Vaultwarden Send) can be opened. */
  maxAccess?: unknown;
}

/** Validated password, expiry date (YYYY-MM-DD) and number of password-link accesses for a new share. */
function shareSettings(body: ShareRequest | undefined) {
  const password = String(body?.password ?? "");
  const weakness = passwordProblem(password);
  if (weakness) throw new HttpError(400, weakness);
  const expireDays = clampInt(body?.expireDays, config.defaultExpireDays, 1, config.maxExpireDays);
  const expireDate = new Date(Date.now() + expireDays * 86_400_000).toISOString().slice(0, 10);
  const maxAccess = clampInt(body?.maxAccess, config.sendMaxAccessCount, 1, config.sendMaxAccessLimit);
  return { password, expireDate, maxAccess };
}

/**
 * Adds a share to the draft's password job (same password, same Send → one password mail) or creates a new job.
 * The password mail only goes out after the mail was really sent, so a job stays open until then.
 * The number of accesses is fixed by the first share of the mail (one Send for all).
 */
async function attachToJob(
  user: User,
  requestedJobId: unknown,
  itemName: string,
  password: string,
  maxAccess: number,
) {
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched) {
    await addFileToJob(id, itemName);
    return id;
  }
  const sendUrl = await createPasswordSend(`CloudShare: ${itemName}`, password, maxAccess);
  return createJob(user.oid, user.upn, itemName, sendUrl, maxAccess);
}

const formatBytes = (bytes: number) =>
  bytes >= 1024 ** 3 ? `${(bytes / 1024 ** 3).toFixed(1)} GB` : `${(bytes / 1024 ** 2).toFixed(1)} MB`;

/**
 * Builds the Fastify app with all routes (no listening, no background jobs) – used by index.ts and by the
 * API tests, which call it with fastify.inject().
 */
/**
 * Rate-limit bucket: the user (oid from the Microsoft token) or, without a token, the client address. The token is
 * only decoded, not verified – a forged oid merely gets its own bucket and is rejected by authenticate() anyway.
 */
function rateLimitKey(req: FastifyRequest): string {
  const header = req.headers.authorization;
  if (header?.startsWith("Bearer ")) {
    try {
      const { oid } = decodeJwt(header.slice("Bearer ".length));
      if (typeof oid === "string" && oid) return `user:${oid}`;
    } catch {
      // not a JWT – fall back to the client address
    }
  }
  return `ip:${req.ip}`;
}

/** Stricter limit for expensive actions (Nextcloud login flow, upload links). */
const strictLimit = { max: Math.max(5, Math.ceil(config.rateLimitPerMinute / 60)), timeWindow: "1 minute" };

export async function buildApp(options: { logger?: boolean } = {}) {
  // Only proxies in private networks (Caddy, an upstream proxy in the LAN) may set X-Forwarded-For – otherwise a
  // client could choose its own address and escape the rate limit.
  const app = Fastify({
    logger: options.logger ?? true,
    bodyLimit: 1024 * 1024,
    trustProxy: ["loopback", "linklocal", "uniquelocal"],
  });

  await app.register(rateLimit, {
    max: config.rateLimitPerMinute,
    timeWindow: "1 minute",
    keyGenerator: rateLimitKey,
    errorResponseBuilder: (_req, context) =>
      Object.assign(
        new HttpError(429, `Zu viele Anfragen – bitte in ${Math.ceil(context.ttl / 1000)} Sekunden erneut versuchen.`, "RATE_LIMITED"),
        { statusCode: 429 },
      ),
  });

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer", bodyLimit: config.chunkSizeBytes + 1024 * 1024 },
    (_req, body, done) => done(null, body),
  );

  app.setErrorHandler((err, req, reply) => {
    if (err instanceof HttpError) return reply.status(err.status).send({ error: err.message, code: err.code });
    req.log.error(err);
    const error = err as Error & { statusCode?: number };
    const status = error.statusCode ?? 500;
    return reply.status(status).send({ error: status < 500 ? error.message : "Interner Fehler" });
  });

  // Add-in files; absent in tests (only the API is exercised there).
  const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../public");
  if (existsSync(publicDir)) await app.register(fastifyStatic, { root: publicDir });

  /**
   * Classic Outlook on Windows runs the event handlers in a JavaScript-only runtime that may only use SSO and
   * network requests if the add-in's origin lists the handler script here. Without it the send handler never
   * reaches the server and Outlook ends with "takes longer than expected" / timeout.
   */
  app.get("/.well-known/microsoft-officeaddins-allowed.json", async () => ({
    allowed: [`${config.publicBaseUrl}/commands.js`],
  }));

  /** Public configuration for the taskpane and the Smart Alert. */
  app.get("/api/config", async () => ({
    thresholdMb: config.thresholdMb,
    attachmentPolicy: config.attachmentPolicy,
    defaultExpireDays: config.defaultExpireDays,
    maxExpireDays: config.maxExpireDays,
    chunkSize: config.chunkSizeBytes,
    sendMaxAccessCount: config.sendMaxAccessCount,
    sendMaxAccessLimit: config.sendMaxAccessLimit,
    mailFooter: config.mailFooter,
  }));

  /** Sender's Microsoft 365 profile – fills the placeholders of MAIL_FOOTER in the taskpane. */
  app.get("/api/profile", async (req) => {
    const user = await authenticate(req);
    const { token } = await onBehalfOf(user, GRAPH_PROFILE_SCOPES);
    return fetchProfile(token);
  });

  // --- Nextcloud connection (Login Flow v2) -------------------------------------

  app.get("/api/nextcloud/status", async (req) => {
    const user = await authenticate(req);
    const creds = getCredentials(user.oid);
    if (!creds) return { connected: false };
    try {
      const { free } = await nextcloudFor(user).userInfo();
      return { connected: true, loginName: creds.loginName, freeBytes: free };
    } catch (err) {
      // App password revoked → nextcloudFor has already removed the connection.
      if (err instanceof HttpError && err.code === NOT_CONNECTED) return { connected: false };
      throw err;
    }
  });

  app.post("/api/nextcloud/connect", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const flow = await startLoginFlow();
    const connectId = randomUUID();
    pendingConnects.set(connectId, { owner: user.oid, flow, expiresAt: Date.now() + CONNECT_TTL });
    return { connectId, loginUrl: flow.loginUrl };
  });

  app.post<{ Params: { id: string } }>("/api/nextcloud/connect/:id/poll", async (req) => {
    const user = await authenticate(req);
    const pending = pendingConnects.get(req.params.id);
    if (!pending || pending.owner !== user.oid || pending.expiresAt < Date.now()) {
      throw new HttpError(410, "Die Anmeldung ist abgelaufen. Bitte erneut verbinden.");
    }
    const creds = await pollLoginFlow(pending.flow);
    if (!creds) return { connected: false };

    pendingConnects.delete(req.params.id);
    // Revoke the old app password if the user was already connected.
    const previous = getCredentials(user.oid);
    if (previous) await new NextcloudClient(previous).revokeAppPassword().catch(() => undefined);
    await saveCredentials(user.oid, creds);
    req.log.info({ user: user.upn, loginName: creds.loginName }, "Nextcloud connected");
    return { connected: true, loginName: creds.loginName };
  });

  app.delete("/api/nextcloud/connection", async (req) => {
    const user = await authenticate(req);
    const creds = getCredentials(user.oid);
    if (creds) await new NextcloudClient(creds).revokeAppPassword().catch(() => undefined);
    await removeCredentials(user.oid);
    return { ok: true };
  });

  // --- Upload and share ----------------------------------------------------------

  app.post<{ Body: { fileName?: string; size?: number } }>("/api/uploads", async (req) => {
    const user = await authenticate(req);
    const fileName = sanitizeFileName(req.body?.fileName);
    const size = Number(req.body?.size);
    const totalChunks = Math.ceil(size / config.chunkSizeBytes);
    if (!Number.isFinite(size) || size <= 0 || totalChunks > 10_000) {
      throw new HttpError(400, "Ungültige Dateigröße");
    }

    const nc = nextcloudFor(user);
    // No fixed maximum: the limit is the user's Nextcloud quota.
    const { id: uid, free } = await nc.userInfo();
    if (free !== null && size > free) {
      throw new HttpError(
        507,
        `Nicht genügend Speicherplatz in Ihrer Nextcloud (frei: ${formatBytes(free)}, benötigt: ${formatBytes(size)}).`,
        QUOTA_EXCEEDED,
      );
    }
    const uploadId = randomUUID();
    const folder = [config.nextcloudFolder, `${new Date().toISOString().slice(0, 10)}_${uploadId.slice(0, 8)}`];
    const target = [...folder, fileName];

    await nc.ensureFolder(uid, folder);
    await nc.startUpload(uid, uploadId, target);

    uploads.set(uploadId, {
      owner: user.oid,
      uid,
      target,
      fileName,
      size,
      totalChunks,
      received: new Set(),
      expiresAt: Date.now() + UPLOAD_TTL,
    });
    return { uploadId, chunkSize: config.chunkSizeBytes, totalChunks };
  });

  app.put<{ Params: { id: string; index: string }; Body: Buffer }>("/api/uploads/:id/chunks/:index", async (req) => {
    const user = await authenticate(req);
    const upload = ownUpload(req.params.id, user);
    const index = Number(req.params.index);
    if (!Number.isInteger(index) || index < 1 || index > upload.totalChunks) {
      throw new HttpError(400, "Ungültige Blocknummer");
    }
    if (!Buffer.isBuffer(req.body) || req.body.length === 0) throw new HttpError(400, "Leerer Block");

    const nc = nextcloudFor(user);
    await nc.putChunk(upload.uid, req.params.id, index, upload.target, req.body);
    upload.received.add(index);
    return { received: upload.received.size, totalChunks: upload.totalChunks };
  });

  app.delete<{ Params: { id: string } }>("/api/uploads/:id", async (req) => {
    const user = await authenticate(req);
    const upload = ownUpload(req.params.id, user);
    const nc = nextcloudFor(user);
    await nc.abortUpload(upload.uid, req.params.id);
    uploads.delete(req.params.id);
    return { ok: true };
  });

  app.post<{ Params: { id: string }; Body: ShareRequest & { jobId?: string } }>(
    "/api/uploads/:id/complete",
    async (req) => {
      const user = await authenticate(req);
      const upload = ownUpload(req.params.id, user);
      if (upload.received.size !== upload.totalChunks) throw new HttpError(409, "Es fehlen noch Blöcke");

      const { password, expireDate, maxAccess } = shareSettings(req.body);

      const nc = nextcloudFor(user);
      await nc.finishUpload(upload.uid, req.params.id, upload.target, upload.size);
      uploads.delete(req.params.id);

      const shareUrl = await nc.createPublicShare(upload.target, password, expireDate);
      const jobId = await attachToJob(user, req.body?.jobId, upload.fileName, password, maxAccess);
      req.log.info({ user: user.upn, file: upload.fileName }, "Share created");
      return { shareUrl, expireDate, jobId, fileName: upload.fileName, size: upload.size };
    },
  );

  // --- Upload links ("Dateien anfordern") ----------------------------------------

  /** Item name of an upload link in the password mail (German, shown to recipients). */
  const UPLOAD_LINK_ITEM = "Upload-Link für Ihre Dateien";

  app.post<{ Body: ShareRequest & { jobId?: string } }>("/api/requests", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const { password, expireDate, maxAccess } = shareSettings(req.body);

    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const folder = [config.requestFolder, `${new Date().toISOString().slice(0, 10)}_${randomUUID().slice(0, 8)}`];
    await nc.ensureFolder(uid, folder);
    // Create-only share ("file drop"): recipients can upload but cannot see or download anything.
    const shareUrl = await nc.createPublicShare(folder, password, expireDate, 4, "Dateianforderung");
    const jobId = await attachToJob(user, req.body?.jobId, UPLOAD_LINK_ITEM, password, maxAccess);
    req.log.info({ user: user.upn, folder: folder.join("/") }, "Upload link created");
    return { shareUrl, expireDate, jobId };
  });

  // --- My shares ("Meine Freigaben") ------------------------------------------------

  const isCloudShareShare = (s: PublicShare) =>
    s.path.startsWith(`/${config.nextcloudFolder}/`) || s.path.startsWith(`/${config.requestFolder}/`);

  /** CloudShare's public links of the user (newest first) with download/upload status. */
  app.get("/api/shares", async (req) => {
    const user = await authenticate(req);
    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const shares = (await nc.listPublicShares())
      .filter(isCloudShareShare)
      .sort((a, b) => b.sharedAt - a.sharedAt)
      .slice(0, 30);

    return Promise.all(
      shares.map(async (s) => {
        const request = s.permissions === 4 && s.isFolder;
        const segments = s.path.split("/").filter(Boolean);
        return {
          id: s.id,
          kind: request ? "request" : "file",
          name: request ? "Dateianforderung" : segments[segments.length - 1],
          url: s.url,
          sharedAt: s.sharedAt,
          expiration: s.expiration,
          downloads: request ? null : await nc.publicDownloads(s.fileId).catch(() => null),
          received: request ? await nc.countFiles(uid, segments).catch(() => null) : null,
        };
      }),
    );
  });

  /** Ends a share ("Zurückziehen"); the file itself stays in the user's Nextcloud. */
  app.delete<{ Params: { id: string } }>("/api/shares/:id", async (req) => {
    const user = await authenticate(req);
    const nc = nextcloudFor(user);
    const share = (await nc.listPublicShares()).find((s) => s.id === req.params.id);
    if (!share || !isCloudShareShare(share)) throw new HttpError(404, "Freigabe nicht gefunden");
    await nc.deleteShare(share.id);
    req.log.info({ user: user.upn, path: share.path }, "Share revoked");
    return { ok: true };
  });

  /**
   * Called by the Smart Alert when the user presses "Send". Nothing is sent here: the job is armed and the watcher
   * (jobs.ts) sends the password mail once the tagged mail shows up in the user's "Sent Items".
   */
  app.post<{ Params: { id: string }; Body: { from?: unknown } }>("/api/jobs/:id/arm", async (req) => {
    const user = await authenticate(req);
    const job = getJob(req.params.id);
    if (!job || job.owner !== user.oid) throw new HttpError(404, "Auftrag nicht gefunden");
    if (job.dispatched) return { ok: true, alreadySent: true };

    // Set by the add-in only if the "From" field differs from the user's own mailbox (shared mailbox, on behalf).
    const from = req.body?.from;
    const mailbox = typeof from === "string" && /^[^\s@/]+@[^\s@/]+$/.test(from) ? from.toLowerCase() : undefined;

    const { token, expiresAt } = await onBehalfOf(user, GRAPH_SCOPES);
    await armJob(req.params.id, { graphToken: token, tokenExpiresAt: expiresAt, armedAt: Date.now(), mailbox });
    req.log.info({ user: user.upn, jobId: req.params.id, mailbox }, "Password mail armed");
    return { ok: true };
  });

  return app;
}
