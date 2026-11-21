import rateLimit from "@fastify/rate-limit";
import fastifyStatic from "@fastify/static";
import Fastify, { type FastifyBaseLogger, type FastifyReply, type FastifyRequest } from "fastify";
import { decodeJwt } from "jose";
import { createHash, randomUUID, timingSafeEqual } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { existsSync } from "node:fs";
import { appGraphToken, authenticate, HttpError, onBehalfOf, type User } from "./auth.js";
import { config, SEND_DAYS_LIMIT } from "./config.js";
import { getCredentials, removeCredentials, saveCredentials } from "./credentials.js";
import { ewsIdFor, fetchContacts, fetchProfile, fetchSensitivityLabels, GRAPH_CONTACT_SCOPES, GRAPH_LABEL_SCOPES, type Contact, GRAPH_PROFILE_SCOPES, GRAPH_SCOPES, sendMail, sendMailAsApp, sentMailState, type MailRecipients, type SensitivityLabel } from "./graph.js";
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
  type Delivery,
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
import { accessIdOf, isMessageToken, messageFor, messagesOf, recordDraftSubject, recordMessage, sentMailFor, updateDraftSubject } from "./history.js";
import { isCloudShareShare, isRequestShare, revokeMessage, revokeShare, tokenOf } from "./shares.js";
import { defaultTemplates, render, smsFiles, templatesFor } from "./templates.js";
import { listContacts, matchContacts } from "./contacts.js";
import { FEEDBACK_MAX_LENGTH, feedbackMail, parseFeedback } from "./feedback.js";
import { clampInt, isEmailAddress } from "./limits.js";
import { formatDate, inRecipientLanguage, inSenderLanguage, knownLanguage, locale, requestLanguageCode, t, tIn, withRequestLanguage } from "./i18n.js";
import { passwordProblem } from "./password.js";
import { signatureHtml } from "./signature.js";
import { availableSettings, defaultSettings, parseSettings, resetSettings, saveSettings, settingsFor } from "./settings.js";
import { CONNECT_TTL, pendingConnects, UPLOAD_TTL, uploads, type UploadState } from "./store.js";
import { addLink, caseById, caseFor, caseForLink, caseOfToken, casesOf, caseSummary, gatewayUrl, linkUrlFor, type ProofCase, type ProofLink } from "./proof.js";
import { caseSubject } from "./proofMail.js";
import { verifiedAddresses } from "./proofPassword.js";
import { confirmFromMail, gatewayCsp, openGateway, passwordAgain, requestCode, verifyCode, type GatewayResult, type Visitor } from "./proofGateway.js";
import { signedProofPdf } from "./proofPdf.js";
import { refreshCase } from "./proofWatcher.js";
import { sign, signedBy } from "./storage.js";
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
  /**
   * "Kennwort senden an" of this link: passwordSms (a number) or passwordTo (an address; "" = the recipients of the
   * mail). Neither sent (clients from before): the mail's one way set by PATCH /api/jobs/:id applies.
   */
  passwordTo?: unknown;
  passwordSms?: unknown;
  /** "Sprache für die Empfänger": language of the password mail and SMS of this link's step (a code of the add-in). */
  language?: unknown;
  /** Subject of the draft the link goes into – "Verlauf" shows it until the mail is sent. */
  subject?: unknown;
  /** "Zustellnachweis" ticked for this link (true), the tracking ID of the mail's case so far, and the client. */
  proof?: unknown;
  proofId?: unknown;
  client?: unknown;
}

/** The way a password goes, as the case records it: "recipients", "mail:<address>", "sms:<masked number>". */
const wayOf = (d: Delivery | undefined) =>
  !d ? [] : d.channel === "recipients" ? ["recipients"] : [`${d.channel}:${d.channel === "sms" ? maskPhone(d.to) : d.to}`];

/** The Send with the password of the step a link joined (its access count is watched for the case). */
function stepSendId(jobId: string | null, token: string): string | undefined {
  const job = jobId ? getJob(jobId) : undefined;
  return job?.separate?.find((part) => part.shares?.some((s) => s.token === token))?.sendId ?? job?.sendId;
}

/**
 * "Zustellnachweis" for a new link, if the sender ticked it (and DELIVERY_PROOF is on): the link joins the mail's case
 * – or opens one – with what went out and how it is protected. Returns the tracking ID and, for a Nextcloud link, the
 * gateway address that goes into the mail instead (the recipient confirms their address there first).
 */
function recordProofLink(
  user: User,
  body: ShareRequest | undefined,
  link: ProofLink & { token: string; jobId: string | null; size?: number; protection: "password" | "none"; delivery?: Delivery; maxAccess?: number },
): { proofId?: string; gateway?: string } {
  if (!config.deliveryProof) return {};
  const id = caseForLink(body?.proof, body?.proofId, user.oid, user.upn, body?.client, {
    recipients: knownLanguage(body?.language),
    sender: requestLanguageCode(),
  });
  if (!id) return {};
  const { token, jobId, size, protection, delivery, maxAccess, ...stored } = link;
  addLink(id, token, { ...stored, createdAt: Date.now() }, protection === "password" ? stepSendId(jobId, token) : undefined, {
    size,
    sha256: stored.sha256,
    protection,
    passwordWays: wayOf(delivery),
    maxAccess,
    language: knownLanguage(body?.language) ?? null,
  });
  return { proofId: id, gateway: stored.url ? gatewayUrl(token) : undefined };
}

/**
 * Links with "Zustellnachweis" and links without one never share a mail: their password would go out together – the
 * password of the proof's files would reach the recipients before they confirmed their address. Which kind the mail's
 * job holds so far: "proof", "plain", or null (no link yet, not the user's, already sent).
 */
function proofKindOf(jobId: unknown, owner: string): "proof" | "plain" | "mixed" | null {
  const job = typeof jobId === "string" ? getJob(jobId) : undefined;
  if (!job || job.owner !== owner || job.dispatched) return null;
  const tokens = (job.separate ?? []).flatMap((part) => (part.shares ?? []).map((s) => s.token));
  if (tokens.length === 0) return null;
  const proof = tokens.map((token) => Boolean(caseOfToken(token)?.links?.[token]));
  return proof.every(Boolean) ? "proof" : proof.some(Boolean) ? "mixed" : "plain";
}

/** Refuses a link that would mix "Zustellnachweis" and no proof in one mail (409 PROOF_MIX). */
function assertNoProofMix(body: ShareRequest & { jobId?: unknown } | undefined, owner: string, withProof: boolean) {
  const kind = proofKindOf(body?.jobId, owner);
  if (kind === null) return;
  if (kind === "mixed" || (kind === "proof") !== withProof) {
    throw new HttpError(409, t(kind === "proof" ? "server.proofMixNeedsProof" : "server.proofMixNoProof"), "PROOF_MIX");
  }
}

/**
 * One password per mail: once the mail's job has a step with a password, every further link with a password must join
 * that step ("Gleiches Kennwort wie zuvor" – the pane offers nothing else). A new step would create a second password
 * and a second password mail – refused before any link is created (409 ONE_PASSWORD). Mails started with one shared
 * password (no `separate`) and sent ones (a new job follows) are not affected.
 */
function assertOnePassword(body: ShareRequest & { jobId?: unknown } | undefined, owner: string, noPassword = false) {
  if (noPassword) return;
  const job = typeof body?.jobId === "string" ? getJob(body.jobId) : undefined;
  if (!job || job.owner !== owner || job.dispatched || !job.separate?.length) return;
  const step = typeof body?.step === "string" ? body.step : undefined;
  if (step && job.separate.some((part) => part.step === step)) return;
  throw new HttpError(409, t("server.onePassword"), "ONE_PASSWORD");
}

/** The way the password of a new link goes out (see ShareRequest), checked like PATCH /api/jobs/:id does. */
function deliveryFrom(body: ShareRequest | undefined): Delivery | undefined {
  if (!body || (body.passwordTo === undefined && body.passwordSms === undefined)) return undefined;
  if (typeof body.passwordSms === "string" && body.passwordSms.trim()) {
    if (!smsEnabled()) throw new HttpError(409, t("server.smsNotConfigured"));
    return { channel: "sms", to: smsNumber(body.passwordSms) };
  }
  const to = typeof body.passwordTo === "string" ? body.passwordTo.trim() : "";
  if (!to) return { channel: "recipients" };
  if (!isEmailAddress(to)) throw new HttpError(400, t("server.recipientsInvalid", { max: 1 }));
  return { channel: "mail", to };
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
  /** The link just created (token, file name, expiry) – for "Verlauf" and download confirmations. */
  share?: StepShare,
  /** "Kennwort senden an" of this link (deliveryFrom); joins the ways of its step. */
  delivery?: Delivery,
  /** Language for the recipients (password mail, SMS) of a new step; a step keeps the one of its first link. */
  language?: string,
) {
  const step = typeof requestedStep === "string" && /^[A-Za-z0-9-]{1,40}$/.test(requestedStep) ? requestedStep : undefined;
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  // Password by SMS: the link of the Send is never sent – its access may last as long as the Send itself.
  const bySms = delivery ? delivery.channel === "sms" : Boolean(existing?.passwordSms);
  if (bySms) link = { ...link, days: config.sendMaxExpireDays };
  const deliveries = delivery ? [delivery] : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched) {
    // A mail started with one shared password (before this version) keeps it.
    if (!existing.separate) {
      await addFileToJob(id, itemName);
      return id;
    }
    if (step && findPart(id, step)) {
      await addFileToPart(id, step, itemName, share, delivery);
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
      deliveries,
      language,
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
      deliveries,
      language,
    });
  }
  return createJob(user.oid, user.upn, itemName, send.url, link.maxAccess, send.id);
}

/**
 * Links "Ohne Kennwort": with the download confirmation switched on they join a job that has no password – it only
 * waits for the sent mail (subject, recipients). Otherwise no job at all (null).
 */
async function attachNoticeJob(user: User, requestedJobId: unknown, share: StepShare, proof = false): Promise<string | null> {
  // Switched off on the server or by the user – unless "Zustellnachweis" needs the sent mail (recipients, Message-ID).
  if (!settingsFor(user.oid).downloadNotice && !(proof && config.deliveryProof)) return null;
  const id = typeof requestedJobId === "string" ? requestedJobId : undefined;
  const existing = id ? getJob(id) : undefined;
  if (id && existing && existing.owner === user.oid && !existing.dispatched && existing.noPassword) {
    await addNoticeShare(id, share);
    return id;
  }
  return createNoticeJob(user.oid, user.upn, share);
}

/**
 * How the password of a link went out, for "Verlauf" (null before sending or for older links).
 * `refused`: seven.io did not take the SMS at all (no message ID: key, credit, daily limit) – as opposed to an SMS
 * reported later as not delivered. Either way nothing went out another way; the sender sends the password again.
 */
function passwordDelivery(token: string, owner: string) {
  const d = deliveryFor(token, owner);
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
 * permission Mail.Send (SMS_NOTICE=true). Without it "Verlauf" shows "nicht zugestellt".
 */
async function noteUndeliveredSms(refs: PasswordRef[], log: Pick<FastifyBaseLogger, "warn">) {
  if (!config.smsNotice || refs.length === 0) return;
  try {
    const token = await appGraphToken();
    for (const ref of refs) {
      if (!settingsFor(ref.owner).smsNotice) continue; // the user switched it off (settings)
      // In the language of the sender's add-in when the mail was sent.
      const { html, subject } = inSenderLanguage(ref.senderLanguage, () => ({
        html: smsProblemMailHtml({ fileNames: ref.fileNames, subject: ref.subject, to: ref.sms ? maskPhone(ref.sms) : "" }),
        subject: t("mail.smsUndeliveredSubject", {
          subject: ref.subject || t("mail.files", { count: ref.fileNames.length, name: ref.fileNames[0] }),
        }),
      }));
      await sendMailAsApp(token, ref.owner, ref.from ?? ref.ownerUpn, subject, html);
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

/** "Verlauf": the draft's subject for a new link (or message) until its mail is sent – see recordDraftSubject. */
function rememberDraftSubject(user: User, body: { subject?: unknown } | undefined, share: StepShare | undefined) {
  if (!share || typeof body?.subject !== "string") return;
  const keepUntil = Date.parse(`${share.expireDate}T00:00:00Z`) + 86_400_000;
  recordDraftSubject(share.token, { owner: user.oid, subject: body.subject.trim().slice(0, 255), keepUntil });
}

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

/** Network errors of fetch (undici puts them in `cause`): the host could not be found or reached. */
const UNREACHABLE_CODES = new Set(["ENOTFOUND", "EAI_AGAIN", "ECONNREFUSED", "ECONNRESET", "ETIMEDOUT", "EHOSTUNREACH", "ENETUNREACH", "UND_ERR_CONNECT_TIMEOUT"]);

/** The text for a service that cannot be reached – Nextcloud, the password service (bw serve) or another host – or null. */
export function unreachableService(err: unknown): string | null {
  let cause: unknown = err;
  for (let depth = 0; cause && depth < 4; depth++) {
    const { code, hostname, address } = cause as { code?: string; hostname?: string; address?: string };
    if (code && UNREACHABLE_CODES.has(code)) {
      const host = hostname ?? address ?? "";
      if (host && host === new URL(config.nextcloudUrl).hostname) return t("server.nextcloudUnreachable", { host });
      if (host === "127.0.0.1" || host === "localhost") return t("server.passwordServiceUnreachable");
      return t("server.serviceUnreachable", { host: host || "?" });
    }
    cause = (cause as { cause?: unknown }).cause;
  }
  return null;
}

/** State of the case a link belongs to – the colour of its seal in "Verlauf" (null: no proof). */
const proofStateOf = (token: string) => {
  const c = caseOfToken(token);
  return c ? caseSummary(c).state : null;
};

/** One case for the list "Zustellnachweise". */
function proofListEntry(c: ProofCase) {
  const sent = [...c.events].reverse().find((e) => e.type === "message.sent")?.data;
  const list = (value: unknown) => (Array.isArray(value) ? value.filter((v): v is string => typeof v === "string") : []);
  const summary = caseSummary(c);
  return {
    id: c.id,
    createdAt: c.createdAt,
    subject: caseSubject(c) || null,
    sentAt: summary.sentAt ?? c.createdAt,
    recipients: [...list(sent?.to), ...list(sent?.cc)],
    files: Object.values(c.links ?? {}).map((l) => l.name),
    downloads: c.events.filter((e) => e.type === "download").length,
    confirmed: c.events.filter((e) => e.type === "recipient.verified").length,
    done: c.done === true,
    state: summary.state,
    fetched: summary.fetched,
    endsOn: summary.endsOn ?? null,
    keptUntil: summary.keptUntil,
    // Like an entry of "Verlauf": the kind symbol, and whether a click on the subject can open the sent mail.
    kind: Object.values(c.links ?? {})[0]?.kind ?? "file",
    mailAvailable: c.tokens.some((tok) => Boolean(sentMailFor(tok, c.owner)?.mailId)),
    // "Kennwort an …": where the password went (addresses, masked numbers); null for links without password.
    passwordTo: Object.values(c.events).some((e) => e.type === "link.created" && e.data.protection !== "none")
      ? [...new Set(c.events.filter((e) => e.type === "password.sent" && typeof e.data.to === "string").map((e) => String(e.data.to)))]
      : null,
  };
}

/** Only mails that really went out: a proof ticked in a draft that was never sent is no proof of anything. */
const sentCasesOf = (owner: string) => casesOf(owner).filter((c) => c.events.some((e) => e.type === "message.sent"));


export async function buildApp(options: { logger?: boolean } = {}) {
  // Only proxies in private networks (Caddy, an upstream proxy in the LAN) may set X-Forwarded-For – otherwise a
  // client could choose its own address and escape the rate limit.
  const app = Fastify({
    logger: options.logger ?? true,
    bodyLimit: 1024 * 1024,
    trustProxy: ["loopback", "linklocal", "uniquelocal"],
  });

  // Error messages in the language of the add-in (header from addin/src/shared/office.ts, api) – before the rate
  // limit, whose message is one of them.
  app.addHook("onRequest", (req, _reply, done) => withRequestLanguage(req.headers["x-cloudshare-language"], done));

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
    // A service CloudShare talks to cannot be reached (wrong address, DNS, network): say which one instead of
    // "Interner Fehler", so the admin knows where to look.
    const unreachable = unreachableService(err);
    if (unreachable) return reply.status(502).send({ error: unreachable, code: "SERVICE_UNREACHABLE" });
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
    // Inside the organisation (INTERNAL_DOMAINS; empty: the user's own domain) a size of its own.
    thresholdInternalMb: config.thresholdInternalMb,
    internalDomains: config.internalDomains,
    attachmentPolicy: config.attachmentPolicy,
    attachmentPolicyInternal: config.attachmentPolicyInternal,
    defaultExpireDays: config.defaultExpireDays,
    maxExpireDays: config.maxExpireDays,
    // Links with password last at most as long as their password in Vaultwarden.
    passwordLinkMaxDays: passwordLinkMaxDays(),
    chunkSize: config.chunkSizeBytes,
    sendMaxAccessCount: config.sendMaxAccessCount,
    sendMaxAccessLimit: config.sendMaxAccessLimit,
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
    // Messages in the pane disappear after this many seconds (0: they stay; progress always stays).
    successMessageSeconds: config.successMessageSeconds,
    // "Nextcloud öffnen" in the account menu.
    nextcloudUrl: config.nextcloudUrl,
    // "Zustellnachweis" in "Gültigkeit und Sprache" (DELIVERY_PROOF).
    deliveryProof: config.deliveryProof,
    // "Feedback senden" in the account menu (FEEDBACK_MAIL; the address itself stays on the server).
    feedback: Boolean(config.feedbackMail),
    // Sensitivity labels of shared files (LABEL_POLICY, LABEL_ENCRYPTED).
    labelPolicy: config.labelPolicy,
    labelEncrypted: config.labelEncrypted,
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

  /**
   * "Feedback senden": a bug report or an idea, mailed from the user's own mailbox to FEEDBACK_MAIL (feedback.ts) in
   * the organisation's language. 404 without FEEDBACK_MAIL. The text is not logged.
   */
  app.post("/api/feedback", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    if (!config.feedbackMail) throw new HttpError(404, t("server.feedbackOff"), "FEEDBACK_OFF");
    const feedback = parseFeedback(req.body);
    if (!feedback) throw new HttpError(400, t("server.feedbackInvalid", { max: FEEDBACK_MAX_LENGTH }));
    const name = await onBehalfOf(user, GRAPH_PROFILE_SCOPES)
      .then(({ token }) => fetchProfile(token))
      .then((profile) => profile.displayName || user.upn)
      .catch(() => user.upn);
    const mail = inRecipientLanguage(config.language, () => feedbackMail(feedback, { name, address: user.upn }));
    const { token } = await onBehalfOf(user, GRAPH_SCOPES);
    await sendMail(token, { to: [config.feedbackMail], cc: [], bcc: [] }, mail.subject, mail.html);
    req.log.info({ user: user.upn, kind: feedback.kind, details: Boolean(feedback.diagnostics) }, "Feedback sent");
    return { ok: true };
  });

  // --- Personal settings (⚙ in the add-in, settings.ts) ------------------------

  /** The settings that apply, the defaults of the environment ("Zurücksetzen") and what the server offers at all. */
  const settingsResponse = (oid: string) => ({ settings: settingsFor(oid), defaults: defaultSettings(), available: availableSettings() });

  /**
   * The sensitivity labels published to the user (Purview, via Graph with InformationProtectionPolicy.Read): name in
   * the add-in's language and priority – the pane names a shared file's label and finds the higher one. `labels: null`
   * when the permission is not consented (or Graph fails): the pane then uses Outlook's own catalog. Kept an hour per
   * user and language.
   */
  const labelCache = new Map<string, { at: number; labels: SensitivityLabel[] | null }>();
  app.get("/api/labels", async (req) => {
    const user = await authenticate(req);
    const language = requestLanguageCode() ?? config.language;
    const key = `${user.oid}|${language}`;
    const cached = labelCache.get(key);
    if (cached && Date.now() - cached.at < 60 * 60_000) return { labels: cached.labels };
    const labels = await onBehalfOf(user, GRAPH_LABEL_SCOPES)
      .then(({ token }) => fetchSensitivityLabels(token, language))
      .catch((err) => {
        req.log.info({ user: user.upn, err: (err as Error).message }, "Sensitivity labels not readable via Graph");
        return null;
      });
    if (labelCache.size > 5000) labelCache.clear();
    labelCache.set(key, { at: Date.now(), labels });
    return { labels };
  });

  /**
   * Suggestions from the user's own Outlook contacts (Graph, optional delegated Contacts.Read) for "Kennwort an":
   * `kind=mail` matches names and addresses, `kind=sms` names and the digits of the numbers; `all=1` lists every entry
   * of that kind A to Z for the "Adressbuch" page. `contacts: null` when the
   * permission is not consented (or Graph fails): the pane then stops asking. The contacts are kept in memory only, ten
   * minutes per user – never stored, the query is not logged.
   */
  const contactCache = new Map<string, { at: number; contacts: Contact[] | null }>();
  app.get("/api/contacts", async (req) => {
    const user = await authenticate(req);
    const { q, kind, all } = req.query as { q?: string; kind?: string; all?: string };
    const cached = contactCache.get(user.oid);
    let contacts = cached && Date.now() - cached.at < 10 * 60_000 ? cached.contacts : undefined;
    if (contacts === undefined) {
      contacts = await onBehalfOf(user, GRAPH_CONTACT_SCOPES)
        .then(({ token }) => fetchContacts(token))
        .catch((err) => {
          req.log.info({ user: user.upn, err: (err as Error).message }, "Contacts not readable via Graph");
          return null;
        });
      if (contactCache.size > 1000) contactCache.clear();
      contactCache.set(user.oid, { at: Date.now(), contacts });
    }
    if (contacts === null) return { contacts: null };
    // The address book page: all of them, A to Z (`all=1`).
    if (all === "1") return { contacts: listContacts(contacts, kind === "sms" ? "sms" : "mail") };
    return { contacts: matchContacts(contacts, String(q ?? "").slice(0, 100), kind === "sms" ? "sms" : "mail", config.smsDefaultCountry) };
  });

  app.get("/api/settings", async (req) => settingsResponse((await authenticate(req)).oid));

  app.put("/api/settings", async (req) => {
    const user = await authenticate(req);
    saveSettings(user.oid, parseSettings(req.body));
    req.log.info({ user: user.upn }, "Settings saved");
    return settingsResponse(user.oid);
  });

  app.delete("/api/settings", async (req) => {
    const user = await authenticate(req);
    resetSettings(user.oid);
    req.log.info({ user: user.upn }, "Settings reset");
    return settingsResponse(user.oid);
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
      hash: createHash("sha256"),
      hashed: 0,
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
    // Fingerprint for "Zustellnachweis": the client sends the chunks in order; a repeated chunk changes nothing.
    if (index === upload.hashed + 1) {
      upload.hash.update(req.body);
      upload.hashed = index;
    }
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

      assertNoProofMix(req.body, user.oid, config.deliveryProof && req.body?.proof === true);
      // "Ohne Kennwort": plain public link – no Vaultwarden Send, no password job, no password mail.
      const noPassword = req.body?.noPassword === true;
      assertOnePassword(req.body, user.oid, noPassword);
      const { password, expireDate, passwordLink } = shareSettings(req.body, !noPassword);
      const delivery = noPassword ? undefined : deliveryFrom(req.body);

      // The fingerprint of the content – complete only if every chunk arrived in order.
      const sha256 = upload.hashed === upload.totalChunks ? (upload.sha256 ??= upload.hash.digest("hex")) : undefined;
      const nc = nextcloudFor(user);
      await nc.finishUpload(upload.uid, req.params.id, upload.target, upload.size);
      uploads.delete(req.params.id);

      const shareUrl = await nc.createPublicShare(upload.target, noPassword ? undefined : password, expireDate);
      await registerForCleanup(user.oid, upload.target, expireDate);
      const share = shareInfo(shareUrl, upload.fileName, expireDate);
      rememberDraftSubject(user, req.body, share);
      const jobId = noPassword
        ? await attachNoticeJob(user, req.body?.jobId, share, req.body?.proof === true)
        : await attachToJob(user, req.body?.jobId, upload.fileName, password, passwordLink, req.body?.step, share, delivery, knownLanguage(req.body?.language));
      const proof = recordProofLink(user, req.body, {
        token: share.token, jobId, url: shareUrl, kind: "file", name: upload.fileName, expireDate, size: upload.size, sha256,
        protection: noPassword ? "none" : "password", delivery, maxAccess: noPassword ? undefined : passwordLink.maxAccess,
      });
      req.log.info({ user: user.upn, file: upload.fileName, noPassword, proof: proof.proofId }, "Share created");
      return { shareUrl: proof.gateway ?? shareUrl, expireDate, jobId, noPassword, fileName: upload.fileName, size: upload.size,
        proofId: proof.proofId ?? null };
    },
  );

  // --- Upload links ("Dateien anfordern") ----------------------------------------

  /** Item name of an upload link in the password mail (German, shown to recipients). */
  /** Name of an upload link in the password mail – in the recipients' language. */
  const uploadLinkItem = (lang?: string) => tIn(lang, "server.requestItem");

  app.post<{ Body: ShareRequest & { jobId?: string } }>("/api/requests", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    assertNoProofMix(req.body, user.oid, false); // an upload link never has a proof
    assertOnePassword(req.body, user.oid);
    const { password, expireDate, passwordLink } = shareSettings(req.body);
    const delivery = deliveryFrom(req.body);

    const nc = nextcloudFor(user);
    const { id: uid } = await nc.userInfo();
    const folder = [config.requestFolder, `${new Date().toISOString().slice(0, 10)}_${randomUUID().slice(0, 8)}`];
    await nc.ensureFolder(uid, folder);
    // Create-only share ("file drop"): recipients can upload but cannot see or download anything.
    const shareUrl = await nc.createPublicShare(folder, password, expireDate, 4, t("server.requestLabel"));
    const language = knownLanguage(req.body?.language);
    const share = shareInfo(shareUrl, uploadLinkItem(), expireDate, []); // a new folder: nothing received yet
    rememberDraftSubject(user, req.body, share);
    const jobId = await attachToJob(user, req.body?.jobId, uploadLinkItem(language), password, passwordLink, req.body?.step,
      share, delivery, language);
    const proof = recordProofLink(user, req.body, {
      token: share.token, jobId, url: shareUrl, kind: "request", name: uploadLinkItem(), expireDate, knownFiles: [],
      protection: "password", delivery, maxAccess: passwordLink.maxAccess,
    });
    req.log.info({ user: user.upn, folder: folder.join("/"), proof: proof.proofId }, "Upload link created");
    return { shareUrl: proof.gateway ?? shareUrl, expireDate, jobId, proofId: proof.proofId ?? null };
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
    assertNoProofMix(req.body, user.oid, false); // a confidential message has no gateway
    assertOnePassword(req.body, user.oid);
    const { password, passwordLink } = shareSettings(req.body);
    const delivery = deliveryFrom(req.body);
    const days = clampInt(req.body?.messageDays, config.sendExpireDays, 1, config.sendMaxExpireDays);
    const expiresAt = new Date(Date.now() + days * 86_400_000);
    const expireDate = expiresAt.toISOString().slice(0, 10);
    // Named in the password mail – in the recipients' language.
    const item = tIn(knownLanguage(req.body?.language), "server.messageItem");
    const send = await createMessageSend(item, text, password, expiresAt);
    // The password is needed exactly as long as the message can be read.
    const link = { ...passwordLink, days: Math.min(passwordLink.days, days), keepUntil: expiresAt.getTime() };
    // "Verlauf" lists it like a link (history.ts) – without key or text – so its password can be shown and sent again.
    const share: StepShare | undefined = send.id
      ? {
          token: recordMessage({ owner: user.oid, sendId: send.id, accessId: accessIdOf(send.url), createdAt: Date.now(), expiresAt: expiresAt.getTime() }),
          name: item,
          expireDate,
          message: true,
        }
      : undefined;
    rememberDraftSubject(user, req.body, share);
    const jobId = await attachToJob(user, req.body?.jobId, item, password, link, req.body?.step, share, delivery, knownLanguage(req.body?.language));
    // No gateway for a message: its link holds the key, which CloudShare does not keep – openings are counted instead.
    const proof: { proofId?: string } = share
      ? recordProofLink(user, req.body, {
          token: share.token, jobId, kind: "message", name: item, expireDate, protection: "password", delivery,
          maxAccess: link.maxAccess,
        })
      : {};
    req.log.info({ user: user.upn, days, length: text.length, proof: proof.proofId }, "Confidential message created");
    // expiresAt: the exact end of the message (its Send), shown with the time in the card.
    return { shareUrl: send.url, expireDate, expiresAt: expiresAt.getTime(), jobId, proofId: proof.proofId ?? null };
  });

  // --- History ("Verlauf") ------------------------------------------------

  /** Upper bound per request: every file link costs one activity lookup in Nextcloud. */
  const SHARES_MAX = 100;

  /**
   * "Verlauf": all valid CloudShare links (newest first) with download/upload status. Expired links are left
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
    /** The mail a link went out with: its subject (title in "Verlauf") and ID (the pane opens it). */
    const mailOf = (token: string) => {
      const mail = sentMailFor(token, user.oid);
      return {
        subject: mail ? mail.subject : null,
        mailId: mail?.mailId ?? null,
        // When and to whom it went (tooltip of the subject) – only known once the sent mail was found.
        sentAt: mail?.sentAt ?? null,
        sentTo: mail?.to ? { to: mail.to, cc: mail.cc ?? [] } : null,
      };
    };

    const links = await Promise.all(
      shares.map(async (s) => {
        const request = isRequestShare(s);
        const segments = s.path.split("/").filter(Boolean);
        const token = tokenOf(s.url);
        const downloads = request ? null : await nc.publicDownloads(s.fileId).catch(() => null);
        return {
          id: s.id,
          kind: request ? "request" : "file",
          name: request ? t("server.requestLabel") : segments[segments.length - 1],
          // A link with "Zustellnachweis" is copied as its gateway address, like it went into the mail.
          url: linkUrlFor(token, s.url),
          // Tracking ID of its case ("Nachweis" in the list), if it has one.
          proofId: caseOfToken(token)?.id ?? null,
          proofState: proofStateOf(token),
          sharedAt: s.sharedAt,
          expiration: s.expiration,
          protected: s.hasPassword,
          // How the password went out (after sending): channel, masked recipient, SMS delivery status.
          password: passwordDelivery(token, user.oid),
          // Its password can still be shown and sent again ("Kennwort" in the list).
          passwordAvailable: passwordRefFor(token, user.oid) !== null,
          // The mail with this link has not been seen sent yet (password not sent) – e.g. a discarded draft – and
          // when CloudShare revokes the link if it stays that way (null: not automatically – also once it was
          // downloaded: then it reached someone, e.g. sent from an Outlook without CloudShare; revokeUnsentLinks).
          unsent: unsent.has(token),
          revokeAt: downloads?.count ? null : (unsent.get(token) ?? null),
          downloads,
          received: request ? await nc.countFiles(uid, segments).catch(() => null) : null,
          ...mailOf(token),
        };
      }),
    );
    // Confidential messages (Vaultwarden Sends, history.ts): no link to copy (CloudShare does not keep the key), no
    // download count (Vaultwarden does not report it) – but the password and "Zurückziehen". `accessId` finds them in
    // a mail being read.
    const messages = messagesOf(user.oid).map((m) => ({
      id: m.token,
      kind: "message",
      name: t("server.messageItem"),
      url: null,
      accessId: m.accessId,
      proofId: caseOfToken(m.token)?.id ?? null,
      proofState: proofStateOf(m.token),
      sharedAt: m.createdAt,
      expiration: new Date(m.expiresAt).toISOString().slice(0, 10),
      protected: true,
      password: passwordDelivery(m.token, user.oid),
      passwordAvailable: passwordRefFor(m.token, user.oid) !== null,
      unsent: unsent.has(m.token),
      revokeAt: unsent.get(m.token) ?? null,
      downloads: null,
      received: null,
      ...mailOf(m.token),
    }));
    return [...links, ...messages].sort((a, b) => b.sharedAt - a.sharedAt).slice(0, SHARES_MAX);
  });

  /**
   * Ends a share ("Zurückziehen"); the file itself stays in the user's Nextcloud. A confidential message ("msg-…") is
   * deleted in Vaultwarden – its link stops working at once.
   */
  app.delete<{ Params: { id: string } }>("/api/shares/:id", async (req) => {
    const user = await authenticate(req);
    if (isMessageToken(req.params.id)) {
      if (!(await revokeMessage(user.oid, req.params.id, req.log))) throw new HttpError(404, t("server.shareNotFound"));
      req.log.info({ user: user.upn }, "Confidential message revoked");
      return { ok: true, message: true };
    }
    const nc = nextcloudFor(user);
    const shares = await nc.listPublicShares();
    const share = shares.find((s) => s.id === req.params.id);
    if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));
    const file = await revokeShare(nc, user.oid, share, shares, req.log);
    req.log.info({ user: user.upn, path: share.path, file }, "Share revoked");
    return { ok: true, file, deleteAfterDays: config.cleanupAfterDays };
  });

  /**
   * "Verlauf": the mail an entry went out with, as the EWS ID Outlook opens (click on the subject). The immutable Graph
   * ID kept when the mail was sent is translated now – Outlook cannot open the immutable form, and the current ID
   * follows a mail that was moved. 404 if no mail is known (sent before this version, shared mailbox) or it is gone.
   */
  app.get<{ Params: { id: string } }>("/api/shares/:id/mail", async (req) => {
    const user = await authenticate(req);
    let token: string;
    if (isMessageToken(req.params.id)) {
      token = req.params.id;
    } else {
      const share = (await nextcloudFor(user).listPublicShares()).find((s) => s.id === req.params.id);
      if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));
      token = tokenOf(share.url);
    }
    return sentMailOf(user, token, req.log);
  });

  /**
   * "Zustellnachweise": the mail of a case, opened like in "Verlauf" (click on the subject) – the mail any of its links
   * went out with.
   */
  app.get<{ Params: { id: string } }>("/api/proofs/:id/mail", async (req) => {
    const user = await authenticate(req);
    const c = caseFor(req.params.id, user.oid);
    if (!c) throw new HttpError(404, t("server.proofNotFound"));
    const token = c.tokens.find((tok) => sentMailFor(tok, user.oid)?.mailId);
    if (!token) throw new HttpError(404, t("server.mailUnknown"));
    return sentMailOf(user, token, req.log);
  });

  /** The EWS ID of the mail a link (token) went out with – 404 if unknown or deleted. */
  async function sentMailOf(user: User, token: string, log: FastifyBaseLogger): Promise<{ ewsId: string }> {
    const mailId = sentMailFor(token, user.oid)?.mailId;
    if (!mailId) throw new HttpError(404, t("server.mailUnknown"));
    // Deleted (also after emptying "Deleted Items": Exchange keeps it hidden for a while) – not opened. If the check
    // itself fails, the mail is opened as before.
    const state = await onBehalfOf(user, GRAPH_SCOPES)
      .then(({ token: mailToken }) => sentMailState(mailToken, mailId))
      .catch((err) => {
        log.warn({ err, user: user.upn }, "State of the sent mail not checked");
        return "ok" as const;
      });
    if (state !== "ok") throw new HttpError(404, t("server.mailDeleted"));
    const { token: graphToken } = await onBehalfOf(user, GRAPH_PROFILE_SCOPES);
    const ewsId = await ewsIdFor(graphToken, mailId).catch((err) => {
      log.warn({ err, user: user.upn }, "Mail ID not translated");
      return null;
    });
    if (!ewsId) throw new HttpError(404, t("server.mailUnknown"));
    return { ewsId };
  }

  /** The user's CloudShare link (or message) with this ID and the entry of its password (404 if it has none any more). */
  async function ownPasswordRef(user: User, id: string): Promise<PasswordRef> {
    let token: string;
    if (isMessageToken(id)) {
      if (!messageFor(id, user.oid)) throw new HttpError(404, t("server.shareNotFound"));
      token = id;
    } else {
      const share = (await nextcloudFor(user).listPublicShares()).find((s) => s.id === id);
      if (!share || !isCloudShareShare(share)) throw new HttpError(404, t("server.shareNotFound"));
      token = tokenOf(share.url);
    }
    const ref = passwordRefFor(token, user.oid);
    if (!ref) throw new HttpError(404, t("server.passwordUnavailable"));
    return ref;
  }

  /**
   * "Kennwort" in "Verlauf": the password of a sent link, read from its Send – only for its sender and only
   * until the link expires – plus where it went the first time, for sending it again the same way.
   */
  app.get<{ Params: { id: string } }>("/api/shares/:id/password", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const ref = await ownPasswordRef(user, req.params.id);
    // "Zustellnachweis": the password is never shown, not even to its sender – the proof says it reached only
    // confirmed addresses (proofPassword.ts). Only where it may go again: the confirmed recipients and the way chosen
    // the first time (the same other address) and the SMS number.
    if (ref.deferred) {
      const verified = verifiedAddresses(caseOfToken(ref.tokens[0] ?? ""), ref.tokens);
      return { deferred: true, verified, recipients: originalRecipients(ref), otherTo: ref.otherTo ?? [], sms: ref.sms ?? "" };
    }
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
      const templates = inRecipientLanguage(ref.language, () => templatesFor({ upn: ref.ownerUpn, department: ref.department }));
      const channel = req.body?.channel;
      // "Zustellnachweis": the password of this step is held back until a recipient confirmed their address with the
      // code (proofPassword.ts). Sending it again must not get round that: nothing before the first confirmation, then
      // only to the confirmed recipients the sender picks (`to`: a list, each in a mail of its own) or the way chosen
      // the first time – the same other address – or by SMS to any number (a corrected one too: SMS numbers are never confirmed).
      const verified = ref.deferred ? verifiedAddresses(caseOfToken(ref.tokens[0] ?? ""), ref.tokens) : [];
      if (ref.deferred && verified.length === 0) throw new HttpError(409, t("server.passwordHeld"), "PASSWORD_HELD");
      const picked = Array.isArray(req.body?.to)
        ? req.body.to.filter((a): a is string => typeof a === "string").map((a) => a.trim().toLowerCase())
        : [];
      if (ref.deferred) {
        const allowed = (channel === "recipients" ? verified : channel === "mail" ? (ref.otherTo ?? []) : []).map((a) => a.toLowerCase());
        const ok = channel === "sms" || (picked.length > 0 && picked.every((a) => allowed.includes(a)));
        if (!ok) throw new HttpError(409, t("server.proofResendConfirmed"), "PROOF_RESEND");
      }

      if (channel === "sms") {
        if (!smsEnabled()) throw new HttpError(409, t("server.smsNotConfigured"));
        const number = smsNumber(target);
        if (!takeSmsQuota(user.oid)) throw new HttpError(429, t("server.smsDailyLimit", { max: config.smsMaxPerDay }));
        const password = await readSendPassword(ref.sendId);
        // In the language the recipients got the password in the first time.
        const text = inRecipientLanguage(ref.language, () =>
          render(templates.sms_template, { password, sender_name: ref.senderName ?? "", files: smsFiles(ref.fileNames ?? []) }));
        const messageId = await sendSms(number, text);
        await recordDelivery(ref.tokens, { owner: user.oid, channel: "sms", to: maskPhone(number), at: Date.now(), messageId, status: "submitted" });
        await updatePasswordRef(ref.sendId, { sms: number });
        req.log.info({ user: user.upn, to: maskPhone(number), messageId }, "Password sent again by SMS");
        return { channel: "sms", to: maskPhone(number) };
      }

      if (channel !== "mail" && channel !== "recipients") throw new HttpError(400, t("server.channelInvalid"));
      // Who gets a mail: the original distribution in one mail, or the one other address (like "Kennwort an" when sending).
      let mails: MailRecipients[];
      let others: string[] = [];
      if (ref.deferred) {
        mails = picked.map((address) => ({ to: [address], cc: [], bcc: [] }));
        if (channel === "mail") others = ref.otherTo ?? []; // kept as it was – "Verlauf" shows the way of the first time
      } else if (channel === "recipients") {
        const original = originalRecipients(ref);
        const all = [...original.to, ...original.cc, ...original.bcc];
        if (all.length === 0) throw new HttpError(400, t("server.recipientsInvalid", { max: 20 }));
        if (picked.length > 0) {
          // The recipients the sender ticked – only addresses the message went to, each in a mail of its own (nobody
          // sees the others, a Bcc stays hidden).
          const known = all.map((address) => address.toLowerCase());
          if (!picked.every((address) => known.includes(address))) throw new HttpError(400, t("server.resendNotRecipient"));
          mails = picked.map((address) => ({ to: [address], cc: [], bcc: [] }));
        } else {
          mails = [original]; // the whole distribution as it was (an older add-in sends no list)
        }
      } else {
        // One address only – several recipients get it with "Empfänger der Nachricht" (an older add-in may send a list).
        const address = target.trim();
        if (!isEmailAddress(address)) throw new HttpError(400, t("server.resendOneAddress"));
        others = [address];
        mails = [{ to: [address], cc: [], bcc: [] }];
      }

      const password = await readSendPassword(ref.sendId);
      const send = await createPasswordSend(`CloudShare: ${ref.fileNames[0]}`, password, ref.maxAccessCount, config.sendExpireDays, ref.keepUntil);
      // The password mail in the language the recipients got it in the first time.
      const { subject, customIntro } = inRecipientLanguage(ref.language, () => ({
        subject: render(templates.mail_password_subject, {
          files: t("mail.files", { count: ref.fileNames.length, name: ref.fileNames[0] }),
          subject: ref.subject || ref.fileNames[0],
          sender_name: ref.senderName ?? "",
        }),
        customIntro: templates.mail_password_intro !== defaultTemplates().mail_password_intro ? templates.mail_password_intro : undefined,
      }));
      // The signature as it is now (settings, profile) – the user sends this mail now.
      const profile = await onBehalfOf(user, GRAPH_PROFILE_SCOPES)
        .then(({ token }) => fetchProfile(token))
        .catch(() => ({ email: user.upn, displayName: ref.senderName ?? "" }));
      const html = inRecipientLanguage(ref.language, () =>
        passwordMailHtml({ fileNames: ref.fileNames, sendUrl: send.url, maxAccessCount: ref.maxAccessCount, subject: ref.subject,
          expireDays: config.sendExpireDays, introTemplate: customIntro, senderName: ref.senderName, expiresOn: formatDate(send.expiresAt),
          signature: signatureHtml(settingsFor(user.oid).signature, profile),
          trackingId: caseOfToken(ref.tokens[0] ?? "")?.id }));
      const { token } = await onBehalfOf(user, GRAPH_SCOPES);
      for (const recipients of mails) await sendMail(token, recipients, subject, html, ref.from);

      // Shown in "Verlauf" like the first time: To and Cc (Bcc stays hidden), or the other addresses.
      const shown = (ref.deferred || (channel === "recipients" && picked.length > 0)
        ? picked
        : channel === "recipients" ? [...mails[0].to, ...mails[0].cc] : others).join(", ");
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
      assertNoProofMix(req.body, user.oid, !request && config.deliveryProof && req.body?.proof === true);
      const noPassword = req.body?.noPassword === true;
      assertOnePassword(req.body, user.oid, noPassword);
      if (request && noPassword) throw new HttpError(400, t("server.requestNeedsPassword"));
      const { password, expireDate, passwordLink } = shareSettings(req.body, !noPassword);
      const delivery = noPassword ? undefined : deliveryFrom(req.body);

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
      rememberDraftSubject(user, req.body, info);
      const jobId = noPassword
        ? await attachNoticeJob(user, req.body?.jobId, info, req.body?.proof === true)
        : await attachToJob(user, req.body?.jobId, request ? uploadLinkItem(knownLanguage(req.body?.language)) : fileName, password, passwordLink, req.body?.step, info, delivery, knownLanguage(req.body?.language));
      // The fingerprint of a file linked again: the one recorded when it was uploaded, if a case knows it.
      const earlier = caseOfToken(tokenOf(share.url))?.links?.[tokenOf(share.url)]?.sha256;
      const proof = recordProofLink(user, req.body, {
        token: info.token, jobId, url: shareUrl, kind: request ? "request" : "file", name: request ? uploadLinkItem() : fileName,
        expireDate, size: request ? undefined : (share.size ?? undefined), sha256: request ? undefined : earlier, knownFiles,
        protection: noPassword ? "none" : "password", delivery, maxAccess: noPassword ? undefined : passwordLink.maxAccess,
      });
      req.log.info({ user: user.upn, path: share.path, noPassword, proof: proof.proofId }, "Share linked again");
      return { kind: request ? "request" : "file", shareUrl: proof.gateway ?? shareUrl, expireDate, jobId, noPassword, fileName,
        size: share.size, proofId: proof.proofId ?? null };
    },
  );

  /**
   * Changes a mail's password job until the password mail went out:
   *   maxAccess  – "Kennwort abrufbar": updates the Send (only jobs that know their Send, created by this version on)
   *   passwordTo – "Kennwort an andere Adresse": the password mail goes only to this address; "" = the recipients
   *   passwordSms – "Kennwort per SMS": the password goes by SMS to this number instead; "" = no SMS
   */
  app.patch<{ Params: { id: string }; Body: { maxAccess?: unknown; passwordTo?: unknown; passwordSms?: unknown; subject?: unknown } }>(
    "/api/jobs/:id",
    async (req) => {
      const user = await authenticate(req);
      const job = getJob(req.params.id);
      if (!job || job.owner !== user.oid) throw new HttpError(404, t("server.jobNotFound"));
      if (job.dispatched) throw new HttpError(409, t("server.passwordAlreadySent"));
      const body = req.body ?? {};

      // The draft's subject changed ("Verlauf" shows it until the mail is sent – also if the draft is discarded).
      if (typeof body.subject === "string") {
        const shares = [...(job.separate ?? []).flatMap((part) => part.shares ?? []), ...(job.noticeShares ?? [])];
        const links = shares.map((s) => ({ token: s.token, keepUntil: Date.parse(`${s.expireDate}T00:00:00Z`) + 86_400_000 }));
        updateDraftSubject(links, user.oid, body.subject.trim().slice(0, 255));
      }

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
    // Before sending: a mail that mixes links with and without "Zustellnachweis" is held back (the send handler shows
    // the reason) – the proof's password would otherwise go out with the others.
    if (proofKindOf(req.params.id, user.oid) === "mixed") throw new HttpError(409, t("server.proofMixSend"), "PROOF_MIX");

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

    // Name and department of the sender for the templates (SMS text, password mail, download confirmation) and the
    // signature of the password mail (own one from the settings or MAIL_FOOTER, filled from the profile). Without the
    // profile: the default template and the signature with what is known (placeholders of the profile dropped).
    await onBehalfOf(user, GRAPH_PROFILE_SCOPES)
      .then(({ token }) => fetchProfile(token))
      .then((profile) => setJobSender(req.params.id, profile.displayName, profile.department, signatureHtml(settingsFor(user.oid).signature, profile)))
      .catch(async (err) => {
        req.log.warn({ err, user: user.upn }, "Profile for templates not readable");
        await setJobSender(req.params.id, "", "", signatureHtml(settingsFor(user.oid).signature, { email: user.upn })).catch(() => undefined);
      });

    const { token, expiresAt } = await onBehalfOf(user, GRAPH_SCOPES);
    await armJob(req.params.id, { graphToken: token, tokenExpiresAt: expiresAt, armedAt: Date.now(), mailbox, language: requestLanguageCode() });
    req.log.info({ user: user.upn, jobId: req.params.id, mailbox }, "Password mail armed");
    return { ok: true };
  });

  // --- "Zustellnachweis" ---------------------------------------------------------

  /** The recipient's browser for the record: its address (behind Caddy) and its user agent. */
  const visitorOf = (req: FastifyRequest): Visitor => ({ ip: req.ip, userAgent: String(req.headers["user-agent"] ?? "") });

  /** A gateway page, or on to the Nextcloud link (with the cookie of a confirmed recipient). */
  const sendGateway = (reply: FastifyReply, result: GatewayResult) => {
    reply.header("Cache-Control", "no-store").header("Content-Security-Policy", gatewayCsp());
    if (result.cookie) reply.header("Set-Cookie", result.cookie);
    if (result.redirect) return reply.redirect(result.redirect, 303);
    return reply.status(result.status).type("text/html; charset=utf-8").send(result.html);
  };

  /**
   * The link in a mail with "Zustellnachweis": the recipient confirms their address with a code (proofGateway.ts) and
   * then goes on to the Nextcloud link. No login – the code is the check.
   */
  app.get<{ Params: { token: string } }>("/r/s/:token", async (req, reply) =>
    sendGateway(reply, openGateway(req.params.token, req.headers.cookie, visitorOf(req))));

  // "Neuen Kennwort-Link anfordern" in the password mail: the address form even for a confirmed browser – a new
  // confirmation with the code sends a fresh password link (proofPassword.ts).
  app.get<{ Params: { token: string } }>("/r/s/:token/password", async (req, reply) =>
    sendGateway(reply, passwordAgain(req.params.token)));

  app.get<{ Params: { token: string }; Querystring: { address?: string; code?: string } }>("/r/s/:token/confirm", async (req, reply) =>
    sendGateway(reply, confirmFromMail(req.params.token, req.query.address, req.query.code)));

  app.post<{ Params: { token: string }; Body: { address?: unknown } }>(
    "/r/s/:token/code",
    { config: { rateLimit: strictLimit } },
    async (req, reply) => sendGateway(reply, await requestCode(req.params.token, req.body?.address, visitorOf(req))),
  );

  app.post<{ Params: { token: string }; Body: { address?: unknown; code?: unknown } }>(
    "/r/s/:token/verify",
    { config: { rateLimit: strictLimit } },
    async (req, reply) => sendGateway(reply, await verifyCode(req.params.token, req.body?.address, req.body?.code, visitorOf(req))),
  );

  /**
   * "Zustellnachweise" (account menu): every case of the user, newest first – also those whose links expired long ago
   * (kept PROOF_RETENTION_MONTHS), so a proof can be found when "Verlauf" no longer lists its link. The pane searches
   * in it (tracking ID, subject, recipients, file names); the seal opens the PDF via /api/proofs/:id/link.
   */
  app.get("/api/proofs", async (req) => {
    const user = await authenticate(req);
    return sentCasesOf(user.oid).slice(0, 1000).map(proofListEntry);
  });


  /** How long the address of a proof PDF works – it is opened in the default browser right away. */
  const PDF_LINK_TTL = 5 * 60_000;
  const pdfSigned = (id: string, until: string, lang: string) => `proof-pdf|${id}|${until}|${lang}`;

  /**
   * "Nachweis" in "Verlauf": a short-lived signed address of the case's PDF, which the pane opens in the default
   * browser (a download from the taskpane does not work in every Outlook). In the language of the add-in.
   */
  app.post<{ Params: { id: string } }>("/api/proofs/:id/link", { config: { rateLimit: strictLimit } }, async (req) => {
    const user = await authenticate(req);
    const c = caseFor(req.params.id, user.oid);
    if (!c) throw new HttpError(404, t("server.proofNotFound"));
    const until = String(Date.now() + PDF_LINK_TTL);
    const lang = requestLanguageCode() ?? "";
    const query = new URLSearchParams({ until, lang, sig: sign(pdfSigned(c.id, until, lang)) });
    return { url: `${config.publicBaseUrl}/proof/${encodeURIComponent(c.id)}.pdf?${query}` };
  });

  app.get<{ Params: { file: string }; Querystring: { until?: string; lang?: string; sig?: string } }>(
    "/proof/:file",
    async (req, reply) => {
      const id = req.params.file.replace(/\.pdf$/, "");
      const { until = "", lang = "", sig } = req.query;
      // The signature binds ID, end and language: only the sender's pane can have asked for this address.
      if (!signedBy(pdfSigned(id, until, lang), sig) || Number(until) < Date.now()) {
        return reply.status(403).type("text/plain; charset=utf-8").send(t("server.proofLinkExpired"));
      }
      const found = caseById(id);
      if (!found) return reply.status(404).type("text/plain; charset=utf-8").send(t("server.proofNotFound"));
      // Up to date: downloads and openings since the last pass of the watcher (every 10 minutes).
      await refreshCase(found, req.log);
      const pdf = await signedProofPdf(found, knownLanguage(lang), req.log);
      const name = `${tIn(knownLanguage(lang), "mail.proof.fileName")}-${found.id}.pdf`;
      req.log.info({ user: found.ownerUpn, case: found.id }, "Delivery proof PDF downloaded");
      return reply
        .header("Cache-Control", "no-store")
        .header("Content-Disposition", `attachment; filename="${name.replace(/[^\w.-]/g, "_")}"`)
        .type("application/pdf")
        .send(pdf);
    },
  );

  return app;
}
