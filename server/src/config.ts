import { callingCodeOf, parseCountries } from "./countries.js";

function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Environment variable ${name} is missing`);
  return value;
}

function number(name: string, fallback: number, allowZero = false): number {
  const value = Number(process.env[name] ?? fallback);
  if (!Number.isFinite(value) || value < 0 || (value === 0 && !allowZero)) throw new Error(`Environment variable ${name} is invalid`);
  return value;
}

/** Lifetime of a Vaultwarden Send in days: whole number from 1 to SEND_DAYS_LIMIT. */
export const SEND_DAYS_LIMIT = 30;
function sendDays(name: string, fallback: number): number {
  const value = Number(process.env[name] ?? fallback);
  if (!Number.isInteger(value) || value < 1 || value > SEND_DAYS_LIMIT) {
    throw new Error(`Environment variable ${name} must be a whole number from 1 to ${SEND_DAYS_LIMIT} (Vaultwarden keeps Sends at most 31 days)`);
  }
  return value;
}

/**
 * DATA_RETENTION_DAYS: entries of the database not changed for this many days are deleted by a daily pass (default 365,
 * "off" = never). A safety net – every kind of entry is removed on its own earlier (see db.ts).
 */
function retentionDays(): number | null {
  const raw = (process.env.DATA_RETENTION_DAYS ?? "365").trim().toLowerCase();
  if (raw === "off") return null;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 1) throw new Error("Environment variable DATA_RETENTION_DAYS is invalid (days or \"off\")");
  return value;
}

/** Days after link expiry until CloudShare deletes the uploaded file; "off" = never (CLEANUP_AFTER_DAYS). */
function cleanupDays(): number | null {
  const raw = (process.env.CLEANUP_AFTER_DAYS ?? "0").trim().toLowerCase();
  if (raw === "off") return null;
  const value = Number(raw);
  if (!Number.isInteger(value) || value < 0) throw new Error("Environment variable CLEANUP_AFTER_DAYS is invalid");
  return value;
}

/** An http(s) address from the environment (empty if unset); anything else stops the start, so a typo shows. */
function webUrl(name: string): string {
  const value = (process.env[name] ?? "").trim();
  if (!value) return "";
  if (!/^https?:\/\/[^\s]+$/i.test(value)) throw new Error(`Environment variable ${name} must be an http(s) address`);
  return value;
}

/** A switch from the environment ("true" = on); unset, it takes the value of `fallbackName` (an older switch). */
export function envFlag(name: string, fallbackName?: string, env: NodeJS.ProcessEnv = process.env): boolean {
  return (env[name] ?? (fallbackName ? env[fallbackName] : undefined)) === "true";
}

/** A regular expression from the environment; null if empty. Throws on an invalid one so a typo stops the start. */
export function envRegex(name: string, fallback: string, raw = process.env[name] ?? fallback): RegExp | null {
  const source = raw.trim();
  if (!source) return null;
  try {
    return new RegExp(source);
  } catch (err) {
    throw new Error(`Environment variable ${name} is no valid regular expression: ${(err as Error).message}`);
  }
}

function oneOf<T extends string>(name: string, values: readonly T[], fallback: T): T {
  const value = (process.env[name] ?? fallback) as T;
  if (!values.includes(value)) throw new Error(`Environment variable ${name} must be one of: ${values.join(", ")}`);
  return value;
}

export const ATTACHMENT_POLICIES = ["enforce", "suggest"] as const;
export type AttachmentPolicy = (typeof ATTACHMENT_POLICIES)[number];
export const LABEL_POLICIES = ["inherit", "suggest", "off"] as const;
export const LABEL_ENCRYPTED = ["warn", "block", "allow"] as const;

export const config = {
  port: number("PORT", 3000),
  publicBaseUrl: required("PUBLIC_BASE_URL").replace(/\/$/, ""),
  /** Shown in Nextcloud as the name of the app password (Settings → Security). */
  appName: "CloudShare Outlook",
  /** <Id> from addin/manifest.xml – names the Office.js custom property that tags drafts ("cecp-<id>"). */
  addinId: process.env.ADDIN_ID ?? "8f3c2a71-5d4e-4b9a-9e2f-6c1d7a0b3e54",

  tenantId: required("ENTRA_TENANT_ID"),
  clientId: required("ENTRA_CLIENT_ID"),
  clientSecret: required("ENTRA_CLIENT_SECRET"),
  /**
   * Signing keys of the Microsoft identity platform. Only the CI load test points this at a mock – never set it in
   * production: whoever controls these keys can sign tokens the backend accepts.
   */
  jwksUrl: process.env.ENTRA_JWKS_URL ?? `https://login.microsoftonline.com/${required("ENTRA_TENANT_ID")}/discovery/v2.0/keys`,

  nextcloudUrl: required("NEXTCLOUD_URL").replace(/\/$/, ""),
  nextcloudFolder: process.env.NEXTCLOUD_FOLDER ?? "CloudShare",
  /**
   * Folder for upload links ("Dateien anfordern"); every request gets its own subfolder. Independent of LANGUAGE on
   * purpose: "Verlauf" recognises CloudShare's links by this folder.
   */
  requestFolder: process.env.NEXTCLOUD_REQUEST_FOLDER ?? "CloudShare-File-Requests",
  /** Optional line at the bottom of the card inserted into the mail, e.g. the company name. */
  mailFooter: process.env.MAIL_FOOTER ?? "",
  /**
   * "Feedback senden" in the add-in (feedback.ts): bug reports and ideas go by mail from the user's mailbox to this
   * address – e.g. whoever looks after CloudShare. Empty: the entry is not offered.
   */
  feedbackMail: (process.env.FEEDBACK_MAIL ?? "").trim(),

  /** Persistent storage for the encrypted Nextcloud app passwords. */
  dataDir: process.env.DATA_DIR ?? "/data",
  /** 32 bytes, Base64 (e.g. `openssl rand -base64 32`). Changing it invalidates all connections. */
  credentialKey: required("CREDENTIAL_KEY"),

  /** From this size (per attachment) CloudShare suggests sending a link. No upper limit – the Nextcloud quota applies. */
  // When a recipient is outside the organisation (INTERNAL_DOMAINS). 0: every attachment (texts without a size).
  thresholdMb: number("ATTACHMENT_THRESHOLD_MB", 5, true),
  // When every recipient is internal – also while none is entered yet. Unset: the same as ATTACHMENT_THRESHOLD_MB.
  thresholdInternalMb: number("ATTACHMENT_THRESHOLD_INTERNAL_MB", Number(process.env.ATTACHMENT_THRESHOLD_MB ?? 5), true),
  // Domains that count as internal (and their subdomains), e.g. "example.com, example.de"; empty = the user's own domain.
  internalDomains: (process.env.INTERNAL_DOMAINS ?? "").split(/[\s,;]+/).map((d) => d.trim().toLowerCase().replace(/^@/, "")).filter(Boolean),
  /**
   * enforce: mails with attachments above the threshold can only be sent via CloudShare (no "Send anyway").
   * suggest: info bar and a prompt on send that offers "Send anyway".
   */
  attachmentPolicy: oneOf<AttachmentPolicy>("ATTACHMENT_POLICY", ATTACHMENT_POLICIES, "suggest"),
  // The same for mails whose recipients are all internal (INTERNAL_DOMAINS). Unset: ATTACHMENT_POLICY.
  attachmentPolicyInternal: oneOf<AttachmentPolicy>("ATTACHMENT_POLICY_INTERNAL", ATTACHMENT_POLICIES,
    (process.env.ATTACHMENT_POLICY as AttachmentPolicy | undefined) ?? "suggest"),
  /**
   * Sensitivity labels (Microsoft Purview) of shared files, read by the add-in: the mail takes over the highest label of
   * its files ("inherit"), the pane only offers it ("suggest") or labels are ignored ("off"). LABEL_ENCRYPTED: a file
   * encrypted by its label is shared after a confirmation ("warn"), refused ("block") or shared without a word ("allow").
   */
  labelPolicy: oneOf("LABEL_POLICY", LABEL_POLICIES, "inherit"),
  labelEncrypted: oneOf("LABEL_ENCRYPTED", LABEL_ENCRYPTED, "warn"),
  defaultExpireDays: number("SHARE_EXPIRE_DAYS", 14),
  maxExpireDays: number("SHARE_MAX_EXPIRE_DAYS", 90),

  bwServeUrl: process.env.BW_SERVE_URL ?? "http://127.0.0.1:8087",
  bwPassword: required("BW_PASSWORD"),
  /** Default number of times the password link (Vaultwarden Send) can be opened; the user may change it per mail. */
  sendMaxAccessCount: number("SEND_MAX_ACCESS_COUNT", 3),
  /** Upper limit for the number of accesses a user may choose. */
  sendMaxAccessLimit: number("SEND_MAX_ACCESS_LIMIT", 20),
  /**
   * How long the password link and confidential messages (Vaultwarden Sends) stay available: default and upper limit
   * of what the user chooses in the pane. Vaultwarden refuses Sends that live longer than 31 days.
   */
  sendExpireDays: sendDays("SEND_EXPIRE_DAYS", 5),
  sendMaxExpireDays: sendDays("SEND_MAX_EXPIRE_DAYS", 30),
  /**
   * SEND_SHOW_CREATOR=true: the Send page names the address of CloudShare's Vaultwarden account ("… created this Send")
   * instead of warning that the creator hid it. Off by default (address hidden, the page shows the warning).
   */
  sendShowCreator: envFlag("SEND_SHOW_CREATOR"),
  /**
   * Files in NEXTCLOUD_FOLDER are deleted (Nextcloud trash) this many days after their link expired; null = never.
   * Received files (NEXTCLOUD_REQUEST_FOLDER) are never deleted automatically.
   */
  cleanupAfterDays: cleanupDays(),
  dataRetentionDays: retentionDays(),

  /**
   * SMS channel via seven.io ("Kennwort per SMS"): off without SEVEN_API_KEY. The sender is a phone number or up to
   * 11 alphanumeric characters; delivery reports arrive at /webhooks/seven?token=<SEVEN_WEBHOOK_SECRET>.
   */
  sevenApiKey: process.env.SEVEN_API_KEY ?? "",
  sevenSender: process.env.SEVEN_SENDER ?? "CloudShare",
  sevenWebhookSecret: process.env.SEVEN_WEBHOOK_SECRET ?? "",
  sevenUrl: process.env.SEVEN_URL ?? "https://gateway.seven.io/api",
  /** Country code for numbers entered without one (e.g. "0170 …" → +49 170 …). */
  smsDefaultCountry: process.env.SMS_DEFAULT_COUNTRY ?? "49",
  /**
   * Countries SMS may go to (ISO codes, calling codes or names, see countries.ts), e.g. "DE, AT, CH" or "DACH";
   * empty = all. Checked in the taskpane while the number is entered and again before every SMS.
   */
  smsAllowedCountries: parseCountries(process.env.SMS_ALLOWED_COUNTRIES ?? "", process.env.LANGUAGE?.trim().toLowerCase() || "en"),
  /**
   * Regular expressions on the number in E.164 ("+491701234567"), checked after the countries: SMS_ALLOWED_NUMBERS
   * (empty = every number) must match, SMS_BLOCKED_NUMBERS must not. Mobile and landline numbers are allowed; the
   * default blocks German freephone, premium-rate and shared-cost numbers (0800, 0900, 0180 …). Only German mobile
   * numbers: SMS_ALLOWED_NUMBERS='^\+491(5|6|7)\d{8,9}$'.
   */
  smsAllowedNumbers: envRegex("SMS_ALLOWED_NUMBERS", ""),
  smsBlockedNumbers: envRegex("SMS_BLOCKED_NUMBERS", "^\\+49(800|900|180)"),
  /** Upper limit of SMS per user and day (cost and abuse protection). */
  smsMaxPerDay: number("SMS_MAX_PER_DAY", 50),

  /**
   * Mail to the sender on the first download of a link. Needs the application permission Mail.Send (client
   * credentials) – restrict it to the CloudShare users with Exchange RBAC for applications.
   */
  downloadNotice: envFlag("DOWNLOAD_NOTICE"),
  /**
   * Mail to the sender when files arrive through an upload link ("Anfordern"), at most one per check (10 minutes).
   * Same application permission Mail.Send as the download confirmation.
   */
  uploadNotice: envFlag("UPLOAD_NOTICE"),
  /**
   * Mail to the sender when seven.io reports a password SMS as not delivered later (webhook) – no user token exists
   * then, so it needs the same application permission Mail.Send. Unset: follows DOWNLOAD_NOTICE, which switched it
   * before this setting existed. (A note about an SMS seven.io refuses right away needs no switch: it goes out with
   * the sender's own token when the mail is sent.)
   */
  smsNotice: envFlag("SMS_NOTICE", "DOWNLOAD_NOTICE"),
  /**
   * "Zustellnachweis" in the pane (proof.ts): per mail, when the sender ticks it – a case with tracking ID, a hash
   * chained record of sending, password, downloads, a recipient confirming their address with a code before the link
   * opens, and a PDF. The code mail goes out with the application permission Mail.Send, so unset it follows
   * DOWNLOAD_NOTICE.
   */
  deliveryProof: envFlag("DELIVERY_PROOF", "DOWNLOAD_NOTICE"),
  /** How long a case is kept after its last event (days; default three years, the regular limitation period). */
  /** How long a delivery proof is kept after its last event, in months (default 36 – three years, the regular limitation period). */
  proofRetentionMonths: number("PROOF_RETENTION_MONTHS", 36),
  /**
   * Signature of the proof PDF (pdfSign.ts): PEM files of an own certificate and key (e.g. from a CA trusted by Adobe);
   * without them – or while they cannot be used, named in the start log – CloudShare signs with a certificate it made
   * itself (CN PROOF_SIGN_NAME, O PROOF_SIGN_ORGANIZATION). PROOF_TSA_URL:
   * RFC 3161 time stamp authority for the signature ("off" = none) – only a hash leaves the server.
   */
  /**
   * The recipients' pages of the delivery proof (proofGateway.ts): the organisation's name after the CloudShare logo,
   * and its website, imprint and privacy notice in the footer – each optional, only http(s) addresses.
   */
  companyName: process.env.COMPANY_NAME?.trim() ?? "",
  companyWebsiteUrl: webUrl("COMPANY_WEBSITE_URL"),
  imprintUrl: webUrl("IMPRINT_URL"),
  privacyUrl: webUrl("PRIVACY_URL"),
  proofSignCertFile: process.env.PROOF_SIGN_CERT_FILE ?? "",
  proofSignKeyFile: process.env.PROOF_SIGN_KEY_FILE ?? "",
  proofSignName: process.env.PROOF_SIGN_NAME?.trim() || "CloudShare",
  proofSignOrganization: process.env.PROOF_SIGN_ORGANIZATION?.trim() ?? "",
  proofTsaUrl: (() => {
    const url = (process.env.PROOF_TSA_URL ?? "http://timestamp.digicert.com").trim();
    return url.toLowerCase() === "off" ? "" : url;
  })(),

  /**
   * Defaults of the personal settings (⚙ in the add-in, settings.ts) – each user may change them for themselves:
   * notices (only where the switch above is on), which ways to send the password the pane offers and which one it
   * preselects, and the language of the add-in ("auto" = as in Outlook). The organisation's signature is MAIL_FOOTER.
   */
  userDefaults: {
    downloadNotice: process.env.DEFAULT_DOWNLOAD_NOTICE !== "false",
    uploadNotice: process.env.DEFAULT_UPLOAD_NOTICE !== "false",
    smsNotice: process.env.DEFAULT_SMS_NOTICE !== "false",
    proofReminder: process.env.DEFAULT_PROOF_REMINDER !== "false",
    channels: (process.env.DEFAULT_PASSWORD_CHANNELS ?? "recipients,mail,sms")
      .split(",")
      .map((c) => c.trim().toLowerCase())
      .filter(Boolean),
    defaultChannel: (process.env.DEFAULT_PASSWORD_CHANNEL ?? "recipients").trim().toLowerCase(),
    language: (process.env.DEFAULT_LANGUAGE ?? "auto").trim().toLowerCase() || "auto",
  },

  /** Optional text templates (SMS, password mail, download confirmation): default → department → user. */
  templatesFile: process.env.TEMPLATES_FILE ?? "/config/templates.json",
  /**
   * Language of everything users and recipients see: locales/<LANGUAGE>.json (see i18n.ts, which reads LANGUAGE
   * itself). Dates in mails are shown in this time zone (the container's TZ).
   */
  language: (process.env.LANGUAGE ?? "en").trim().toLowerCase() || "en",
  timeZone: process.env.TZ || "Europe/Berlin",

  /**
   * Requests per minute and user (oid from the Microsoft token) or, without one, per client address. Generous on
   * purpose – a 5 GB upload is about 500 chunk requests. Expensive actions (Nextcloud login, upload links) get 1/60.
   */
  rateLimitPerMinute: number("RATE_LIMIT_PER_MINUTE", 600),
  /** Diagnostics: the Smart Alert reports its steps to POST /api/trace (logged). Off by default. */
  clientTrace: process.env.CLIENT_TRACE === "true",
  /**
   * Messages in the pane disappear after this many seconds (a ring around the tick, "!" or "i" runs down; it stops
   * while the pointer is on the message). 0 = they stay. Progress and a failed start stay.
   */
  successMessageSeconds: (() => {
    const value = Number(process.env.SUCCESS_MESSAGE_SECONDS ?? 5);
    if (!Number.isFinite(value) || value < 0 || value > 600) {
      throw new Error("Environment variable SUCCESS_MESSAGE_SECONDS must be a number of seconds from 0 to 600");
    }
    return value;
  })(),

  /** Nextcloud chunking v2: at least 5 MiB per chunk (except the last), at most 10,000 chunks. */
  chunkSizeBytes: 10 * 1024 * 1024,
};

if (config.feedbackMail && !/^[^\s@,;<>"()]+@[^\s@,;<>"()]+\.[^\s@,;<>"()]{2,}$/.test(config.feedbackMail)) {
  throw new Error("FEEDBACK_MAIL must be one email address");
}

if (config.sendExpireDays > config.sendMaxExpireDays) {
  throw new Error("SEND_EXPIRE_DAYS must not be larger than SEND_MAX_EXPIRE_DAYS");
}

const allowedCodes = config.smsAllowedCountries.map(callingCodeOf);
if (allowedCodes.length > 0 && !allowedCodes.includes(config.smsDefaultCountry)) {
  throw new Error(
    `SMS_DEFAULT_COUNTRY (${config.smsDefaultCountry}) must belong to SMS_ALLOWED_COUNTRIES – numbers like 0170 … would all be refused`,
  );
}

// Entries of links that are still valid must not be deleted: "Verlauf" needs them until the links expire.
if (config.dataRetentionDays !== null && config.dataRetentionDays < config.maxExpireDays) {
  throw new Error(
    `DATA_RETENTION_DAYS (${config.dataRetentionDays}) must be at least SHARE_MAX_EXPIRE_DAYS (${config.maxExpireDays}) – ` +
      "otherwise data of links that are still valid would be deleted",
  );
}
