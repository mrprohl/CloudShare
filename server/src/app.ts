import rateLimit from "@fastify/rate-limit";
import fastifyStatic from "@fastify/static";
import Fastify, { type FastifyBaseLogger, type FastifyRequest } from "fastify";
import { decodeJwt } from "jose";
import { randomUUID, timingSafeEqual } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { appGraphToken, authenticate, HttpError, onBehalfOf, type User } from "./auth.js";
import { config, SEND_DAYS_LIMIT } from "./config.js";
import { getCredentials, removeCredentials, saveCredentials } from "./credentials.js";
import { fetchProfile, GRAPH_PROFILE_SCOPES, GRAPH_SCOPES, sendMail, sendMailAsApp, type MailRecipients } from "./graph.js";
import {
  addFileToJob,
  addNoticeShare,
  addSeparatePart,
  armJob,
  createJob,
  createNoticeJob,
  createSeparateJob,
  findPart,
  addFileToPart,
  getJob,
  setJobMaxAccess,
  setJobPasswordTo,
  setJobSender,
  setPartsExpiry,
  takeSmsQuota,
  unsentShares,
  type StepShare,
} from "./jobs.js";
import {
  NextcloudClient,
  NOT_CONNECTED,
  pollLoginFlow,
  QUOTA_EXCEEDED,
  startLoginFlow,
} from "./nextcloud.js";
import { registerForCleanup } from "./cleanup.js";
import { deliveryFor, recordDelivery, updateSmsStatus } from "./deliveries.js";
import {
  originalRecipients,
  passwordRefFor,
  passwordRefsForTokens,
  updatePasswordRef,
  type PasswordRef,
} from "./passwordRefs.js";
import { passwordMailHtml, smsProblemMailHtml } from "./passwordMail.js";
import {
  mapReportStatus,
  maskPhone,
  normalizePhone,
  parseDeliveryReport,
  sendSms,
  smsCountryCodes,
  smsCountryList,
  smsEnabled,
  smsNumber,
  smsNumberProblem,
} from "./sms.js";
import { isCloudShareShare, isRequestShare, revokeShare, tokenOf } from "./shares.js";
import { defaultTemplates, render, templatesFor } from "./templates.js";
import { clampInt, isEmailAddress } from "./limits.js";
import { formatDate, locale, t } from "./i18n.js";
import { passwordProblem } from "./password.js";
import { CONNECT_TTL, pendingConnects, UPLOAD_TTL, uploads, type UploadState } from "./store.js";
import {
  createMessageSend,
  createPasswordSend,
  MESSAGE_MAX_LENGTH,
  readSendPassword,
  setSendExpiry,
  setSendMaxAccess,
} from "./vaultwarden.js";

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
  /** How many days the password link stays available ("Kennwort-Link gültig für"). */
  passwordDays?: unknown;
  /** ID of the "Als Link einfügen" step (one password and one password mail per step, see attachToJob). */
  step?: unknown;
}

/** The password link of a step: how often and for how many days it can be opened, and how long its Send lives. */
interface PasswordLink {
  maxAccess: number;
  days: number;
  /** Expiry of the links (ms): the Send – and so the password for "Kennwort erneut senden" – lives until then. */
  keepUntil: number;
}

/**
 * Longest validity of a link with password: it lasts only as long as its password in Vaultwarden, and Vaultwarden
 * keeps a Send at most 31 days. Links "Ohne Kennwort" may last up to SHARE_MAX_EXPIRE_DAYS.
 */
const passwordLinkMaxDays = () => Math.min(config.maxExpireDays, SEND_DAYS_LIMIT);

/**
 * Validated password, expiry date (YYYY-MM-DD) and password link (accesses, days, lifetime of the Send) for a new
 * share. `withPassword: false` (file links "Ohne Kennwort") skips the password check; the password is then ignored.
 */
function shareSettings(body: ShareRequest | undefined, withPassword = true) {
  const password = withPassword ? String(body?.password ?? "") : "";
  const weakness = withPassword ? passwordProblem(password) : undefined;
  if (weakness) throw new HttpError(400, weakness);
  const maxDays = withPassword ? passwordLinkMaxDays() : config.maxExpireDays;
  const expireDays = clampInt(body?.expireDays, Math.min(config.defaultExpireDays, maxDays), 1, maxDays);
  const keepUntil = Date.now() + expireDays * 86_400_000;
  const expireDate = new Date(keepUntil).toISOString().slice(0, 10);
  const maxAccess = clampInt(body?.maxAccess, config.sendMaxAccessCount, 1, config.sendMaxAccessLimit);
  // The recipients' access to the password ends at the latest with the link.
  const days = Math.min(clampInt(body?.passwordDays, config.sendExpireDays, 1, config.sendMaxExpireDays), expireDays);
  return { password, expireDate, passwordLink: { maxAccess, days, keepUntil } satisfies PasswordLink };
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
  link: PasswordLink,
  requestedStep: unknown,
  /** The link just created (token, file name, expiry) – for "Meine Freigaben" and download confirmations. */
  share?: StepShare,
) {
  const step = typeof requestedStep === "string" && /^[A-Za-z0-9-]{1,40}$/.test(requestedStep) ? requestedStep : undefined;
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched) {
    // Password by SMS: the link of the Send is never sent – its access may last as long as the Send itself.
    if (existing.passwordSms) link = { ...link, days: config.sendMaxExpireDays };
    // A mail started with one shared password (before this version) keeps it.
    if (!existing.separate) {
      await addFileToJob(id, itemName);
      return id;
    }
    if (step && findPart(id, step)) {
      await addFileToPart(id, step, itemName, share);
      return id;
    }
    const send = await createPasswordSend(`CloudShare: ${itemName}`, password, link.maxAccess, link.days, link.keepUntil);
    await addSeparatePart(id, {
      step: step ?? randomUUID(),
      fileNames: [itemName],
      sendUrl: send.url,
      sendId: send.id,
      maxAccessCount: link.maxAccess,
      sendExpiresAt: send.expiresAt,
      sendKeepUntil: link.keepUntil,
      shares: share ? [share] : undefined,
    });
    return id;
  }
  const send = await createPasswordSend(`CloudShare: ${itemName}`, password, link.maxAccess, link.days, link.keepUntil);
  if (step) {
    return createSeparateJob(user.oid, user.upn, {
      step,
      fileNames: [itemName],
      sendUrl: send.url,
      sendId: send.id,
      maxAccessCount: link.maxAccess,
      sendExpiresAt: send.expiresAt,
      sendKeepUntil: link.keepUntil,
      shares: share ? [share] : undefined,
    });
  }
  return createJob(user.oid, user.upn, itemName, send.url, link.maxAccess, send.id);
}

/**
 * Links "Ohne Kennwort": with the download confirmation switched on they join a job that has no password – it only
 * waits for the sent mail (subject, recipients). Otherwise no job at all (null).
 */
async function attachNoticeJob(user: User, requestedJobId: unknown, share: StepShare): Promise<string | null> {
  if (!config.downloadNotice) return null;
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched && existing.noPassword) {
    await addNoticeShare(id, share);
    return id;
  }
  return createNoticeJob(user.oid, user.upn, share);
}

/**
 * How the password of a link went out, for "Meine Freigaben" (null before sending or for older links).
 * `refused`: seven.io did not take the SMS at all (no message ID: key, credit, daily limit) – as opposed to an SMS
 * reported later as not delivered. Either way nothing went out another way; the sender sends the password again.
 */
function passwordDelivery(url: string, owner: string) {
  const d = deliveryFor(tokenOf(url), owner);
  return d
    ? {
        channel: d.channel,
        to: d.to,
        status: d.status ?? null,
        at: d.at,
        deliveredAt: d.deliveredAt ?? null,
        refused: d.channel === "sms" && d.status === "failed" && !d.messageId,
      }
    : null;
}

/**
 * seven.io reported an SMS as not delivered: tell the sender – without password or link – so they send it again in
 * CloudShare. No user token exists at that point; only CloudShare's own app token can send, which needs the application
 * permission Mail.Send (SMS_NOTICE=true). Without it "Meine Freigaben" shows "nicht zugestellt".
 */
async function noteUndeliveredSms(refs: PasswordRef[], log: Pick<FastifyBaseLogger, "warn">) {
  if (!config.smsNotice || refs.length === 0) return;
  try {
    const token = await appGraphToken();
    for (const ref of refs) {
      const files = t("mail.files", { count: ref.fileNames.length, name: ref.fileNames[0] });
      const html = smsProblemMailHtml({ fileNames: ref.fileNames, subject: ref.subject, to: ref.sms ? maskPhone(ref.sms) : "" });
      await sendMailAsApp(token, ref.owner, ref.from ?? ref.ownerUpn, t("mail.smsUndeliveredSubject", { subject: ref.subject || files }), html);
    }
  } catch (err) {
    log.warn({ err }, "Note about an undelivered SMS not sent");
  }
}

/**
 * Token (last URL segment), name and expiry of a link just created. Upload links also carry the files already in their
 * folder (`knownFiles`), so the upload notification only reports files that arrive later.
 */
const shareInfo = (url: string, name: string, expireDate: string, knownFiles?: string[]): StepShare => ({
  token: url.replace(/\/+$/, "").split("/").pop() ?? "",
  name,
  expireDate,
  ...(knownFiles ? { request: true as const, knownFiles } : {}),
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
    // Links with password last at most as long as their password in Vaultwarden.
    passwordLinkMaxDays: passwordLinkMaxDays(),
    chunkSize: config.chunkSizeBytes,
    sendMaxAccessCount: config.sendMaxAccessCount,
    sendMaxAccessLimit: config.sendMaxAccessLimit,
    mailFooter: config.mailFooter,
    trace: config.clientTrace,
    smsEnabled: smsEnabled(),
    smsDefaultCountry: config.smsDefaultCountry,
    // SMS_ALLOWED_COUNTRIES: calling codes for the check while typing and the names for the hint ([] / "" = all).
    smsCountryCodes: smsCountryCodes(),
    smsCountries: smsCountryList(),
    // SMS_ALLOWED_NUMBERS / SMS_BLOCKED_NUMBERS as regular expressions on the E.164 number ("" = none).
    smsAllowedNumbers: config.smsAllowedNumbers?.source ?? "",
    smsBlockedNumbers: config.smsBlockedNumbers?.source ?? "",
    messageMaxLength: MESSAGE_MAX_LENGTH,
    sendExpireDays: config.sendExpireDays,
    sendMaxExpireDays: config.sendMaxExpireDays,
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
      const { known, failedTokens } = await updateSmsStatus(report.messageId, mapReportStatus(report.status));
      req.log.info({ messageId: report.messageId, status: report.status, known }, "SMS delivery report");
      if (failedTokens.length > 0) await noteUndeliveredSms(passwordRefsForTokens(failedTokens), req.log);
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
      const { password, expireDate, passwordLink } = shareSettings(req.body, !noPassword);

      const nc = nextcloudFor(user);
      await nc.finishUpload(upload.uid, req.params.id, upload.target, upload.size);
      uploads.delete(req.params.id);

      const shareUrl = await nc.createPublicShare(upload.target, noPassword ? undefined : password, expireDate);
      await registerForCleanup(user.oid, upload.target, expireDate);
      const share = shareInfo(shareUrl, upload.fileName, expireDate);
      const jobId = noPassword
        ? await attachNoticeJob(user, req.body?.jobId, share)
        : await attachToJob(user, req.body?.jobId, upload.fileName, password, passwordLink, req.body?.step, share);
      req.log.info({ user: user.upn, file: upload.fileName, noPassword }, "Share created");
      return { shareUrl, expireDate, jobId, noPassword, fileName: upload.fileName, size: upload.size };
    },
  );

  // --- Upload links ("Dateien anfordern") ----------------------------------------

  /** Item name of an upload link in the password mail (German, shown to recipients). */
  const uploadLinkItem = () => t("server.requestItem");

  app.post<{ Body: ShareRequest & { jobId?: string } }>("/api/requests", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const { password, expireDate, passwordLink } = shareSettings(req.body);

    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const folder = [config.requestFolder, `${new Date().toISOString().slice(0, 10)}_${randomUUID().slice(0, 8)}`];
    await nc.ensureFolder(uid, folder);
    // Create-only share ("file drop"): recipients can upload but cannot see or download anything.
    const shareUrl = await nc.createPublicShare(folder, password, expireDate, 4, t("server.requestLabel"));
    const jobId = await attachToJob(user, req.body?.jobId, uploadLinkItem(), password, passwordLink, req.body?.step,
      shareInfo(shareUrl, uploadLinkItem(), expireDate, [])); // a new folder: nothing received yet
    req.log.info({ user: user.upn, folder: folder.join("/") }, "Upload link created");
    return { shareUrl, expireDate, jobId };
  });

  // --- Confidential message ("Vertrauliche Nachricht") ---------------------------
  /**
   * A text instead of a file: stored only as a password-protected Vaultwarden Send (never in Nextcloud), the link goes
   * into the mail as a card, the password goes out like for files (same job, step, channel). Always with a password;
   * available for `messageDays` ("Nachricht gültig für", 1 to SEND_MAX_EXPIRE_DAYS, default SEND_EXPIRE_DAYS).
   */
  app.post<{ Body: ShareRequest & { jobId?: string; text?: unknown; messageDays?: unknown } }>("/api/messages", async (req) => {
    const user = await authenticate(req);
    const text = typeof req.body?.text === "string" ? req.body.text.replace(/\r\n/g, "\n").trim() : "";
    if (!text || text.length > MESSAGE_MAX_LENGTH) {
      throw new HttpError(400, t("server.messageInvalid", { max: MESSAGE_MAX_LENGTH }));
    }
    const { password, passwordLink } = shareSettings(req.body);
    const days = clampInt(req.body?.messageDays, config.sendExpireDays, 1, config.sendMaxExpireDays);
    const expiresAt = new Date(Date.now() + days * 86_400_000);
    const expireDate = expiresAt.toISOString().slice(0, 10);
    const item = t("server.messageItem");
    const send = await createMessageSend(item, text, password, expiresAt);
    // The password is needed exactly as long as the message can be read.
    const link = { ...passwordLink, days: Math.min(passwordLink.days, days), keepUntil: expiresAt.getTime() };
    const jobId = await attachToJob(user, req.body?.jobId, item, password, link, req.body?.step);
    req.log.info({ user: user.upn, days, length: text.length }, "Confidential message created");
    return { shareUrl: send.url, expireDate, jobId };
  });

  // --- My shares ("Meine Freigaben") ------------------------------------------------

  /** Upper bound per request: every file link costs one activity lookup in Nextcloud. */
  const SHARES_MAX = 100;

  /**
   * "Meine Freigaben": all valid CloudShare links (newest first) with download/upload status. Expired links are left
   * out – Nextcloud removes them only in its daily housekeeping.
   */
  app.get("/api/shares", async (req) => {
    const user = await authenticate(req);
    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const today = new Date().toISOString().slice(0, 10);
    const shares = (await nc.listPublicShares())
      .filter((s) => isCloudShareShare(s) && (s.expiration === null || s.expiration >= today))
      .sort((a, b) => b.sharedAt - a.sharedAt)
      .slice(0, SHARES_MAX);
    const unsent = unsentShares(user.oid);

    return Promise.all(
      shares.map(async (s) => {
        const request = isRequestShare(s);
        const segments = s.path.split("/").filter(Boolean);
        const token = tokenOf(s.url);
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
          // Its password can still be shown and sent again ("Kennwort" in the list).
          passwordAvailable: passwordRefFor(token, user.oid) !== null,
          // The mail with this link has not been seen sent yet (password not sent) – e.g. a discarded draft – and
          // when CloudShare revokes the link if it stays that way (null: not automatically).
          unsent: unsent.has(token),
          revokeAt: unsent.get(token) ?? null,
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
    const file = await revokeShare(nc, user.oid, share, shares, req.log);
    req.log.info({ user: user.upn, path: share.path, file }, "Share revoked");
    return { ok: true, file, deleteAfterDays: config.cleanupAfterDays };
  });

  /** The user's CloudShare link with this ID and the entry of its password (404 if it has none any more). */
  async function ownPasswordRef(user: User, id: string): Promise<PasswordRef> {
    const share = (await nextcloudFor(user).listPublicShares()).find((s) => s.id === id);
    if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));
    const ref = passwordRefFor(tokenOf(share.url), user.oid);
    if (!ref) throw new HttpError(404, t("server.passwordUnavailable"));
    return ref;
  }

  /**
   * "Kennwort" in "Meine Freigaben": the password of a sent link, read from its Send – only for its sender and only
   * until the link expires – plus where it went the first time, for sending it again the same way.
   */
  app.get<{ Params: { id: string } }>("/api/shares/:id/password", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const ref = await ownPasswordRef(user, req.params.id);
    const password = await readSendPassword(ref.sendId);
    req.log.info({ user: user.upn, shareId: req.params.id }, "Password shown to its sender");
    return { password, recipients: originalRecipients(ref), otherTo: ref.otherTo ?? [], sms: ref.sms ?? "" };
  });

  /**
   * "Kennwort erneut senden" – the same ways as the first time, never as plain text in a mail:
   *   recipients – the password mail with the distribution of the sent mail (To, Cc, Bcc as they were);
   *   mail       – to other addresses, each in its own mail (nobody sees the others);
   *   sms        – by SMS to a number.
   * A mail gets a new Send with the same password (the recipients' access to the first one may have ended), deleted
   * with the links; it goes out as the sender through Graph like the first password mail.
   */
  app.post<{ Params: { id: string }; Body: { channel?: unknown; to?: unknown } }>(
    "/api/shares/:id/password/send",
    { config: { rateLimit: strictLimit } },
    async (req) => {
      const user = await authenticate(req);
      const ref = await ownPasswordRef(user, req.params.id);
      const target = typeof req.body?.to === "string" ? req.body.to : "";
      const templates = templatesFor({ upn: ref.ownerUpn, department: ref.department });
      const channel = req.body?.channel;

      if (channel === "sms") {
        if (!smsEnabled()) throw new HttpError(409, t("server.smsNotConfigured"));
        const number = smsNumber(target);
        if (!takeSmsQuota(user.oid)) throw new HttpError(429, t("server.smsDailyLimit", { max: config.smsMaxPerDay }));
        const password = await readSendPassword(ref.sendId);
        const messageId = await sendSms(number, render(templates.sms_template, { password, sender_name: ref.senderName ?? "" }));
        await recordDelivery(ref.tokens, { owner: user.oid, channel: "sms", to: maskPhone(number), at: Date.now(), messageId, status: "submitted" });
        await updatePasswordRef(ref.sendId, { sms: number });
        req.log.info({ user: user.upn, to: maskPhone(number), messageId }, "Password sent again by SMS");
        return { channel: "sms", to: maskPhone(number) };
      }

      if (channel !== "mail" && channel !== "recipients") throw new HttpError(400, t("server.channelInvalid"));
      // Who gets a mail: the original distribution in one mail, or every other address in a mail of its own.
      let mails: MailRecipients[];
      let others: string[] = [];
      if (channel === "recipients") {
        const original = originalRecipients(ref);
        if (original.to.length + original.cc.length + original.bcc.length === 0) {
          throw new HttpError(400, t("server.recipientsInvalid", { max: 20 }));
        }
        mails = [original];
      } else {
        others = [...new Set(target.split(/[\s,;]+/).filter(Boolean))];
        if (others.length === 0 || others.length > 20 || !others.every(isEmailAddress)) {
          throw new HttpError(400, t("server.recipientsInvalid", { max: 20 }));
        }
        mails = others.map((address) => ({ to: [address], cc: [], bcc: [] }));
      }

      const password = await readSendPassword(ref.sendId);
      const send = await createPasswordSend(`CloudShare: ${ref.fileNames[0]}`, password, ref.maxAccessCount, config.sendExpireDays, ref.keepUntil);
      const files = t("mail.files", { count: ref.fileNames.length, name: ref.fileNames[0] });
      const customIntro = templates.mail_password_intro !== defaultTemplates().mail_password_intro ? templates.mail_password_intro : undefined;
      const subject = render(templates.mail_password_subject, { files, subject: ref.subject || ref.fileNames[0], sender_name: ref.senderName ?? "" });
      const html = passwordMailHtml({ fileNames: ref.fileNames, sendUrl: send.url, maxAccessCount: ref.maxAccessCount, subject: ref.subject,
        expireDays: config.sendExpireDays, introTemplate: customIntro, senderName: ref.senderName, expiresOn: formatDate(send.expiresAt) });
      const { token } = await onBehalfOf(user, GRAPH_SCOPES);
      for (const recipients of mails) await sendMail(token, recipients, subject, html, ref.from);

      // Shown in "Meine Freigaben" like the first time: To and Cc (Bcc stays hidden), or the other addresses.
      const shown = (channel === "recipients" ? [...mails[0].to, ...mails[0].cc] : others).join(", ");
      await recordDelivery(ref.tokens, { owner: user.oid, channel: "mail", to: shown, at: Date.now() });
      await updatePasswordRef(ref.sendId, { otherTo: channel === "mail" ? others : undefined, resentSend: send.id });
      req.log.info({ user: user.upn, channel, mails: mails.length }, "Password sent again by mail");
      return { channel, to: shown };
    },
  );

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

      const request = isRequestShare(share);
      const noPassword = req.body?.noPassword === true;
      if (request && noPassword) throw new HttpError(400, t("server.requestNeedsPassword"));
      const { password, expireDate, passwordLink } = shareSettings(req.body, !noPassword);

      const segments = share.path.split("/").filter(Boolean);
      const fileName = segments[segments.length - 1];
      const shareUrl = request
        ? await nc.createPublicShare(segments, password, expireDate, 4, t("server.requestLabel"))
        : await nc.createPublicShare(segments, noPassword ? undefined : password, expireDate);
      if (!request) await registerForCleanup(user.oid, segments, expireDate);
      // An upload folder linked again may already hold files: the upload notification only reports new ones.
      const knownFiles = request
        ? config.uploadNotice ? ((await nc.listFiles((await nc.userInfo()).id, segments)) ?? []) : []
        : undefined;
      const info = shareInfo(shareUrl, request ? uploadLinkItem() : fileName, expireDate, knownFiles);
      const jobId = noPassword
        ? await attachNoticeJob(user, req.body?.jobId, info)
        : await attachToJob(user, req.body?.jobId, request ? uploadLinkItem() : fileName, password, passwordLink, req.body?.step, info);
      req.log.info({ user: user.upn, path: share.path, noPassword }, "Share linked again");
      return { kind: request ? "request" : "file", shareUrl, expireDate, jobId, noPassword, fileName, size: share.size };
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
        if (!smsEnabled()) throw new HttpError(409, t("server.smsNotConfigured"));
        const number = smsNumber(body.passwordSms);
        await setJobPasswordTo(req.params.id, undefined, number);
        // The link of an SMS step's Send is never sent: let its access last as long as the Send (until the links expire).
        const expiresAt = Date.now() + config.sendMaxExpireDays * 86_400_000;
        for (const part of job.separate ?? []) {
          if (part.dispatched || !part.sendId) continue;
          await setSendExpiry(part.sendId, expiresAt).catch((err) => req.log.warn({ err, jobId: req.params.id }, "Send lifetime not extended"));
        }
        await setPartsExpiry(req.params.id, expiresAt);
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
    // A number no SMS may go to is ignored here: the job keeps the number the taskpane set before.
    if (passwordSms && smsEnabled() && !smsNumberProblem(passwordSms)) {
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
