import rateLimit from "@fastify/rate-limit";
import fastifyStatic from "@fastify/static";
import Fastify, { type FastifyRequest } from "fastify";
import { decodeJwt } from "jose";
import { randomUUID, timingSafeEqual } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { authenticate, HttpError, onBehalfOf, type User } from "./auth.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials, saveCredentials } from "./credentials.js";
import { fetchProfile, GRAPH_PROFILE_SCOPES, GRAPH_SCOPES } from "./graph.js";
import {
  addFileToJob,
  addSeparatePart,
  armJob,
  createJob,
  createSeparateJob,
  findPart,
  addFileToPart,
  getJob,
  setJobMaxAccess,
  setJobPasswordTo,
  setJobSender,
  type StepShare,
} from "./jobs.js";
import {
  NextcloudClient,
  NOT_CONNECTED,
  pollLoginFlow,
  QUOTA_EXCEEDED,
  startLoginFlow,
  type PublicShare,
} from "./nextcloud.js";
import { expireOnRevoke, registerForCleanup } from "./cleanup.js";
import { deliveryFor, updateSmsStatus } from "./deliveries.js";
import { mapReportStatus, maskPhone, normalizePhone, parseDeliveryReport, smsEnabled } from "./sms.js";
import { clampInt, isEmailAddress } from "./limits.js";
import { locale, t } from "./i18n.js";
import { passwordProblem } from "./password.js";
import { CONNECT_TTL, pendingConnects, UPLOAD_TTL, uploads, type UploadState } from "./store.js";
import { createMessageSend, createPasswordSend, MESSAGE_MAX_DAYS, MESSAGE_MAX_LENGTH, setSendMaxAccess } from "./vaultwarden.js";

function nextcloudFor(user: User): NextcloudClient {
  const creds = getCredentials(user.oid);
  if (!creds) throw new HttpError(409, t("server.notConnected"), NOT_CONNECTED);
  return new NextcloudClient(creds, () => removeCredentials(user.oid));
}

function sanitizeFileName(name: unknown): string {
  const cleaned = String(name ?? "")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .trim()
    .slice(0, 200);
  if (!cleaned || cleaned === "." || cleaned === "..") throw new HttpError(400, t("server.invalidFileName"));
  return cleaned;
}

function ownUpload(id: string, user: User): UploadState {
  const upload = uploads.get(id);
  if (!upload || upload.owner !== user.oid) throw new HttpError(404, t("server.uploadNotFound"));
  return upload;
}

interface ShareRequest {
  password?: unknown;
  expireDays?: unknown;
  /** How often the password link (Vaultwarden Send) can be opened. */
  maxAccess?: unknown;
  /** ID of the "Als Link einfügen" step (one password and one password mail per step, see attachToJob). */
  step?: unknown;
}

/**
 * Validated password, expiry date (YYYY-MM-DD) and number of password-link accesses for a new share.
 * `withPassword: false` (file links "Ohne Kennwort") skips the password check; the password is then ignored.
 */
function shareSettings(body: ShareRequest | undefined, withPassword = true) {
  const password = withPassword ? String(body?.password ?? "") : "";
  const weakness = withPassword ? passwordProblem(password) : undefined;
  if (weakness) throw new HttpError(400, weakness);
  const expireDays = clampInt(body?.expireDays, config.defaultExpireDays, 1, config.maxExpireDays);
  const expireDate = new Date(Date.now() + expireDays * 86_400_000).toISOString().slice(0, 10);
  const maxAccess = clampInt(body?.maxAccess, config.sendMaxAccessCount, 1, config.sendMaxAccessLimit);
  return { password, expireDate, maxAccess };
}

/**
 * Adds a share to the draft's password job or creates a new job. The password mail only goes out after the mail was
 * really sent, so a job stays open until then.
 * `step` (one ID per "Als Link einfügen" click, sent by the taskpane): every step gets its own password, its own Send
 * and its own password mail; further files of the same step join it. Without `step` (drafts and cached clients from
 * before) all shares of the mail share one password, one Send and one password mail as before.
 */
async function attachToJob(
  user: User,
  requestedJobId: unknown,
  itemName: string,
  password: string,
  maxAccess: number,
  requestedStep: unknown,
  /** The link just created (token, file name, expiry) – for "Meine Freigaben" and download confirmations. */
  share?: StepShare,
) {
  const step = typeof requestedStep === "string" && /^[A-Za-z0-9-]{1,40}$/.test(requestedStep) ? requestedStep : undefined;
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched) {
    // A mail started with one shared password (before this version) keeps it.
    if (!existing.separate) {
      await addFileToJob(id, itemName);
      return id;
    }
    if (step && findPart(id, step)) {
      await addFileToPart(id, step, itemName, share);
      return id;
    }
    const send = await createPasswordSend(`CloudShare: ${itemName}`, password, maxAccess);
    await addSeparatePart(id, {
      step: step ?? randomUUID(),
      fileNames: [itemName],
      sendUrl: send.url,
      sendId: send.id,
      maxAccessCount: maxAccess,
      shares: share ? [share] : undefined,
    });
    return id;
  }
  const send = await createPasswordSend(`CloudShare: ${itemName}`, password, maxAccess);
  if (step) {
    return createSeparateJob(user.oid, user.upn, {
      step,
      fileNames: [itemName],
      sendUrl: send.url,
      sendId: send.id,
      maxAccessCount: maxAccess,
      shares: share ? [share] : undefined,
    });
  }
  return createJob(user.oid, user.upn, itemName, send.url, maxAccess, send.id);
}

/** How the password of a link went out, for "Meine Freigaben" (null before sending or for older links). */
function passwordDelivery(url: string, owner: string) {
  const d = deliveryFor(url.replace(/\/+$/, "").split("/").pop() ?? "", owner);
  return d ? { channel: d.channel, to: d.to, status: d.status ?? null, at: d.at, deliveredAt: d.deliveredAt ?? null } : null;
}

/** Token (last URL segment), name and expiry of a link just created. */
const shareInfo = (url: string, name: string, expireDate: string): StepShare => ({
  token: url.replace(/\/+$/, "").split("/").pop() ?? "",
  name,
  expireDate,
});

const formatBytes = (bytes: number) =>
  bytes >= 1024 ** 3
    ? `${(bytes / 1024 ** 3).toLocaleString(locale(), { maximumFractionDigits: 1 })} GB`
    : `${(bytes / 1024 ** 2).toLocaleString(locale(), { maximumFractionDigits: 1 })} MB`;

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
        new HttpError(429, t("server.rateLimited", { seconds: Math.ceil(context.ttl / 1000) }), "RATE_LIMITED"),
        { statusCode: 429 },
      ),
  });

  app.addContentTypeParser(
    "application/octet-stream",
    { parseAs: "buffer", bodyLimit: config.chunkSizeBytes + 1024 * 1024 },
    (_req, body, done) => done(null, body),
  );

  // Form-encoded bodies (classic seven.io delivery report callback) – Fastify only parses JSON by default.
  app.addContentTypeParser("application/x-www-form-urlencoded", { parseAs: "string" }, (_req, body, done) => {
    done(null, Object.fromEntries(new URLSearchParams(String(body))));
  });

  app.setErrorHandler((err, req, reply) => {
    if (err instanceof HttpError) return reply.status(err.status).send({ error: err.message, code: err.code });
    req.log.error(err);
    const error = err as Error & { statusCode?: number };
    const status = error.statusCode ?? 500;
    return reply.status(status).send({ error: status < 500 ? error.message : t("server.internalError") });
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
    trace: config.clientTrace,
    smsEnabled: smsEnabled(),
    smsDefaultCountry: config.smsDefaultCountry,
    messageMaxLength: MESSAGE_MAX_LENGTH,
    messageMaxDays: MESSAGE_MAX_DAYS,
  }));

  /**
   * Diagnostics (CLIENT_TRACE=true only): steps of the Smart Alert, to find where sending hangs in a client.
   * No login – the handler may be stuck exactly at the sign-in. Only short, validated fields are logged.
   */
  // GET with query parameters: a "simple" request that a restricted frame still sends (no preflight, no body).
  type TraceFields = { run?: unknown; step?: unknown; ms?: unknown; detail?: unknown; host?: unknown; origin?: unknown };
  app.route<{ Querystring: TraceFields; Body: TraceFields }>({
    method: ["GET", "POST"],
    url: "/api/trace",
    handler: async (req, reply) => {
      if (!config.clientTrace) return reply.status(404).send({ error: "Not Found" });
      const data: TraceFields = (req.method === "GET" ? req.query : req.body) ?? {};
      const text = (value: unknown, pattern: RegExp) => (typeof value === "string" && pattern.test(value) ? value : "?");
      const ms = Number(data.ms);
      req.log.info(
        {
          trace: text(data.run, /^[a-z0-9]{1,12}$/),
          step: text(data.step, /^[a-z0-9:._-]{1,40}$/),
          ms: Number.isFinite(ms) ? Math.round(ms) : null,
          host: text(data.host, /^[A-Za-z0-9 ._-]{1,40}$/),
          origin: text(data.origin, /^(null|none|https?:\/\/[A-Za-z0-9.-]+(:\d{1,5})?)$/),
          detail: typeof data.detail === "string" ? data.detail.slice(0, 300) : undefined,
        },
        "Smart Alert trace",
      );
      return { ok: true };
    },
  });

  /**
   * Delivery reports of seven.io (set as callback URL in the seven.io dashboard:
   * https://<host>/webhooks/seven?token=<SEVEN_WEBHOOK_SECRET>). No login – secured by the secret, compared in
   * constant time. Idempotent: the same report twice changes nothing. Webhook format and classic callback accepted.
   */
  app.route<{ Querystring: Record<string, unknown>; Body: unknown }>({
    method: ["GET", "POST"],
    url: "/webhooks/seven",
    handler: async (req, reply) => {
      const expected = Buffer.from(config.sevenWebhookSecret);
      const given = Buffer.from(typeof req.query.token === "string" ? req.query.token : "");
      if (!expected.length || expected.length !== given.length || !timingSafeEqual(expected, given)) {
        return reply.status(404).send({ error: "Not Found" });
      }
      const report = parseDeliveryReport(req.body, req.query);
      if (!report) return { ok: true, ignored: true }; // other events (e.g. inbound SMS) are not used
      const known = await updateSmsStatus(report.messageId, mapReportStatus(report.status));
      req.log.info({ messageId: report.messageId, status: report.status, known }, "SMS delivery report");
      return { ok: true };
    },
  });

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
      const nc = nextcloudFor(user);
      const [{ free, used, total }, passwordEnforced] = await Promise.all([
        nc.userInfo(),
        // Unknown → offer "Ohne Kennwort"; Nextcloud's own error message still explains a refusal.
        nc.linkPasswordEnforced().catch(() => false),
      ]);
      return {
        connected: true,
        loginName: creds.loginName,
        freeBytes: free,
        usedBytes: used,
        totalBytes: total,
        passwordEnforced,
      };
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
      throw new HttpError(410, t("server.loginExpired"));
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
      throw new HttpError(400, t("server.invalidFileSize"));
    }

    const nc = nextcloudFor(user);
    // No fixed maximum: the limit is the user's Nextcloud quota.
    const { id: uid, free } = await nc.userInfo();
    if (free !== null && size > free) {
      throw new HttpError(
        507,
        t("server.quotaExceededDetail", { free: formatBytes(free), needed: formatBytes(size) }),
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
      throw new HttpError(400, t("server.invalidChunk"));
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

  app.post<{ Params: { id: string }; Body: ShareRequest & { jobId?: string; noPassword?: unknown } }>(
    "/api/uploads/:id/complete",
    async (req) => {
      const user = await authenticate(req);
      const upload = ownUpload(req.params.id, user);
      if (upload.received.size !== upload.totalChunks) throw new HttpError(409, t("server.missingChunks"));

      // "Ohne Kennwort": plain public link – no Vaultwarden Send, no password job, no password mail.
      const noPassword = req.body?.noPassword === true;
      const { password, expireDate, maxAccess } = shareSettings(req.body, !noPassword);

      const nc = nextcloudFor(user);
      await nc.finishUpload(upload.uid, req.params.id, upload.target, upload.size);
      uploads.delete(req.params.id);

      const shareUrl = await nc.createPublicShare(upload.target, noPassword ? undefined : password, expireDate);
      await registerForCleanup(user.oid, upload.target, expireDate);
      const jobId = noPassword ? null : await attachToJob(user, req.body?.jobId, upload.fileName, password, maxAccess, req.body?.step,
        shareInfo(shareUrl, upload.fileName, expireDate));
      req.log.info({ user: user.upn, file: upload.fileName, noPassword }, "Share created");
      return { shareUrl, expireDate, jobId, fileName: upload.fileName, size: upload.size };
    },
  );

  // --- Upload links ("Dateien anfordern") ----------------------------------------

  /** Item name of an upload link in the password mail (German, shown to recipients). */
  const uploadLinkItem = () => t("server.requestItem");

  app.post<{ Body: ShareRequest & { jobId?: string } }>("/api/requests", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const { password, expireDate, maxAccess } = shareSettings(req.body);

    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const folder = [config.requestFolder, `${new Date().toISOString().slice(0, 10)}_${randomUUID().slice(0, 8)}`];
    await nc.ensureFolder(uid, folder);
    // Create-only share ("file drop"): recipients can upload but cannot see or download anything.
    const shareUrl = await nc.createPublicShare(folder, password, expireDate, 4, t("server.requestLabel"));
    const jobId = await attachToJob(user, req.body?.jobId, uploadLinkItem(), password, maxAccess, req.body?.step,
      shareInfo(shareUrl, uploadLinkItem(), expireDate));
    req.log.info({ user: user.upn, folder: folder.join("/") }, "Upload link created");
    return { shareUrl, expireDate, jobId };
  });

  // --- Confidential message ("Vertrauliche Nachricht") ---------------------------
  /**
   * A text instead of a file: stored only as a password-protected Vaultwarden Send (never in Nextcloud), the link goes
   * into the mail as a card, the password goes out like for files (same job, step, channel). Always with a password;
   * at most MESSAGE_MAX_DAYS days (Vaultwarden's limit), even if "Link gültig für" is longer.
   */
  app.post<{ Body: ShareRequest & { jobId?: string; text?: unknown } }>("/api/messages", async (req) => {
    const user = await authenticate(req);
    const text = typeof req.body?.text === "string" ? req.body.text.replace(/\r\n/g, "\n").trim() : "";
    if (!text || text.length > MESSAGE_MAX_LENGTH) {
      throw new HttpError(400, t("server.messageInvalid", { max: MESSAGE_MAX_LENGTH }));
    }
    const { password, maxAccess } = shareSettings(req.body);
    const wanted = clampInt(req.body?.expireDays, config.defaultExpireDays, 1, config.maxExpireDays);
    const days = Math.min(wanted, MESSAGE_MAX_DAYS);
    const expiresAt = new Date(Date.now() + days * 86_400_000);
    const expireDate = expiresAt.toISOString().slice(0, 10);
    const item = t("server.messageItem");
    const send = await createMessageSend(item, text, password, expiresAt);
    const jobId = await attachToJob(user, req.body?.jobId, item, password, maxAccess, req.body?.step);
    req.log.info({ user: user.upn, days, length: text.length }, "Confidential message created");
    return { shareUrl: send.url, expireDate, jobId, capped: days < wanted };
  });

  // --- My shares ("Meine Freigaben") ------------------------------------------------

  const isCloudShareShare = (s: PublicShare) =>
    s.path.startsWith(`/${config.nextcloudFolder}/`) || s.path.startsWith(`/${config.requestFolder}/`);

  /** "Meine Freigaben" shows the links of the last 30 days. */
  const SHARES_DAYS = 30;
  /** Upper bound per request: every file link costs one activity lookup in Nextcloud. */
  const SHARES_MAX = 100;

  /** CloudShare's public links of the last 30 days (newest first) with download/upload status. */
  app.get("/api/shares", async (req) => {
    const user = await authenticate(req);
    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const since = Date.now() - SHARES_DAYS * 86_400_000;
    const shares = (await nc.listPublicShares())
      .filter((s) => isCloudShareShare(s) && s.sharedAt >= since)
      .sort((a, b) => b.sharedAt - a.sharedAt)
      .slice(0, SHARES_MAX);

    return Promise.all(
      shares.map(async (s) => {
        const request = s.permissions === 4 && s.isFolder;
        const segments = s.path.split("/").filter(Boolean);
        return {
          id: s.id,
          kind: request ? "request" : "file",
          name: request ? t("server.requestLabel") : segments[segments.length - 1],
          url: s.url,
          sharedAt: s.sharedAt,
          expiration: s.expiration,
          protected: s.hasPassword,
          // How the password went out (after sending): channel, masked recipient, SMS delivery status.
          password: passwordDelivery(s.url, user.oid),
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
    const shares = await nc.listPublicShares();
    const share = shares.find((s) => s.id === req.params.id);
    if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));
    await nc.deleteShare(share.id);

    // A revoked file link counts as expired: the upload folder goes to the trash (CLEANUP_AFTER_DAYS) unless another
    // valid link still points to it. Received files of an upload link are never deleted.
    const segments = share.path.split("/").filter(Boolean);
    const folder = `/${segments.slice(0, 2).join("/")}/`;
    const today = new Date().toISOString().slice(0, 10);
    const otherLinkActive = shares.some(
      (s) => s.id !== share.id && s.path.startsWith(folder) && (s.expiration === null || s.expiration >= today),
    );
    const request = share.permissions === 4 && share.isFolder;
    const file = request
      ? "kept"
      : await expireOnRevoke(user.oid, segments, otherLinkActive, async (target) => {
          const { id: uid } = await nc.userInfo();
          await nc.deletePath(uid, target);
        });
    req.log.info({ user: user.upn, path: share.path, file }, "Share revoked");
    return { ok: true, file, deleteAfterDays: config.cleanupAfterDays };
  });

  /**
   * "Als Link einfügen": a new public link to the same file or upload folder for the mail being written, with this
   * mail's settings – its password and password mail (joins the draft's job), or no password for files. The original
   * password cannot be recovered (Nextcloud only keeps a hash), and the old link stays valid for its recipients.
   */
  app.post<{ Params: { id: string }; Body: ShareRequest & { jobId?: string; noPassword?: unknown } }>(
    "/api/shares/:id/link",
    { config: { rateLimit: strictLimit } },
    async (req) => {
      const user = await authenticate(req);
      const nc = nextcloudFor(user);
      const share = (await nc.listPublicShares()).find((s) => s.id === req.params.id);
      if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));

      const request = share.permissions === 4 && share.isFolder;
      const noPassword = req.body?.noPassword === true;
      if (request && noPassword) throw new HttpError(400, t("server.requestNeedsPassword"));
      const { password, expireDate, maxAccess } = shareSettings(req.body, !noPassword);

      const segments = share.path.split("/").filter(Boolean);
      const fileName = segments[segments.length - 1];
      const shareUrl = request
        ? await nc.createPublicShare(segments, password, expireDate, 4, t("server.requestLabel"))
        : await nc.createPublicShare(segments, noPassword ? undefined : password, expireDate);
      if (!request) await registerForCleanup(user.oid, segments, expireDate);
      const jobId = noPassword
        ? null
        : await attachToJob(user, req.body?.jobId, request ? uploadLinkItem() : fileName, password, maxAccess, req.body?.step,
            shareInfo(shareUrl, request ? uploadLinkItem() : fileName, expireDate));
      req.log.info({ user: user.upn, path: share.path, noPassword }, "Share linked again");
      return { kind: request ? "request" : "file", shareUrl, expireDate, jobId, fileName, size: share.size };
    },
  );

  /**
   * Changes a mail's password job until the password mail went out:
   *   maxAccess  – "Kennwort abrufbar": updates the Send (only jobs that know their Send, created by this version on)
   *   passwordTo – "Kennwort an andere Adresse": the password mail goes only to this address; "" = the recipients
   *   passwordSms – "Kennwort per SMS": the password goes by SMS to this number instead; "" = no SMS
   */
  app.patch<{ Params: { id: string }; Body: { maxAccess?: unknown; passwordTo?: unknown; passwordSms?: unknown } }>(
    "/api/jobs/:id",
    async (req) => {
      const user = await authenticate(req);
      const job = getJob(req.params.id);
      if (!job || job.owner !== user.oid) throw new HttpError(404, t("server.jobNotFound"));
      if (job.dispatched) throw new HttpError(409, t("server.passwordAlreadySent"));
      const body = req.body ?? {};

      if (body.passwordSms !== undefined && body.passwordSms !== "") {
        const number = typeof body.passwordSms === "string" ? normalizePhone(body.passwordSms) : null;
        if (!smsEnabled()) throw new HttpError(409, t("server.smsNotConfigured"));
        if (!number) throw new HttpError(400, t("server.phoneInvalid"));
        await setJobPasswordTo(req.params.id, undefined, number);
        req.log.info({ user: user.upn, jobId: req.params.id, sms: maskPhone(number) }, "Password channel: SMS");
      } else if (body.passwordTo !== undefined || body.passwordSms === "") {
        const passwordTo = typeof body.passwordTo === "string" ? body.passwordTo.trim() : body.passwordTo === undefined ? "" : null;
        if (passwordTo === null || (passwordTo !== "" && !isEmailAddress(passwordTo))) {
          throw new HttpError(400, t("server.emailInvalid"));
        }
        await setJobPasswordTo(req.params.id, passwordTo || undefined);
        req.log.info({ user: user.upn, jobId: req.params.id, other: Boolean(passwordTo) }, "Password recipient changed");
      }

      if (body.maxAccess !== undefined) {
        if (!job.sendId) {
          throw new HttpError(409, t("server.accessFixed"), "SEND_FIXED");
        }
        const maxAccess = clampInt(body.maxAccess, config.sendMaxAccessCount, 1, config.sendMaxAccessLimit);
        await setSendMaxAccess(job.sendId, maxAccess);
        await setJobMaxAccess(req.params.id, maxAccess);
        req.log.info({ user: user.upn, jobId: req.params.id, maxAccess }, "Password-link accesses changed");
      }
      return {
        maxAccess: job.maxAccessCount,
        passwordTo: job.passwordTo ?? "",
        passwordSms: job.passwordSms ? maskPhone(job.passwordSms) : "",
      };
    },
  );

  /**
   * Called by the Smart Alert when the user presses "Send". Nothing is sent here: the job is armed and the watcher
   * (jobs.ts) sends the password mail once the tagged mail shows up in the user's "Sent Items".
   */
  app.post<{ Params: { id: string }; Body: { from?: unknown; passwordTo?: unknown; passwordSms?: unknown } }>("/api/jobs/:id/arm", async (req) => {
    const user = await authenticate(req);
    const job = getJob(req.params.id);
    if (!job || job.owner !== user.oid) throw new HttpError(404, t("server.jobNotFound"));
    if (job.dispatched) return { ok: true, alreadySent: true };

    // Set by the add-in only if the "From" field differs from the user's own mailbox (shared mailbox, on behalf).
    const from = req.body?.from;
    const mailbox = typeof from === "string" && /^[^\s@/]+@[^\s@/]+$/.test(from) ? from.toLowerCase() : undefined;

    // "Kennwort an andere Adresse" from the draft – normally already set by the taskpane, ignored if not valid.
    const passwordTo = req.body?.passwordTo;
    const passwordSms = typeof req.body?.passwordSms === "string" ? normalizePhone(req.body.passwordSms) : null;
    if (passwordSms && smsEnabled()) {
      await setJobPasswordTo(req.params.id, undefined, passwordSms);
    } else if (typeof passwordTo === "string" && isEmailAddress(passwordTo.trim())) {
      await setJobPasswordTo(req.params.id, passwordTo.trim());
    }

    // Name and department of the sender for the templates (SMS text, password mail, download confirmation). Optional:
    // without them the default template is used.
    await onBehalfOf(user, GRAPH_PROFILE_SCOPES)
      .then(({ token }) => fetchProfile(token))
      .then((profile) => setJobSender(req.params.id, profile.displayName, profile.department))
      .catch((err) => req.log.warn({ err, user: user.upn }, "Profile for templates not readable"));

    const { token, expiresAt } = await onBehalfOf(user, GRAPH_SCOPES);
    await armJob(req.params.id, { graphToken: token, tokenExpiresAt: expiresAt, armedAt: Date.now(), mailbox });
    req.log.info({ user: user.upn, jobId: req.params.id, mailbox }, "Password mail armed");
    return { ok: true };
  });

  return app;
}
