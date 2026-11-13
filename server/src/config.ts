function required(name: string): string {
  const value = process.env[name];
  if (!value) throw new Error(`Environment variable ${name} is missing`);
  return value;
}

function number(name: string, fallback: number): number {
  const value = Number(process.env[name] ?? fallback);
  if (!Number.isFinite(value) || value <= 0) throw new Error(`Environment variable ${name} is invalid`);
  return value;
}

function oneOf<T extends string>(name: string, values: readonly T[], fallback: T): T {
  const value = (process.env[name] ?? fallback) as T;
  if (!values.includes(value)) throw new Error(`Environment variable ${name} must be one of: ${values.join(", ")}`);
  return value;
}

export const ATTACHMENT_POLICIES = ["enforce", "suggest"] as const;
export type AttachmentPolicy = (typeof ATTACHMENT_POLICIES)[number];

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
  nextcloudFolder: process.env.NEXTCLOUD_FOLDER ?? "Outlook-Anhaenge",
  /** Folder for upload links ("Dateien anfordern"); every request gets its own subfolder. */
  requestFolder: process.env.NEXTCLOUD_REQUEST_FOLDER ?? "Outlook-Dateianforderungen",
  /** Optional line at the bottom of the card inserted into the mail, e.g. the company name. */
  mailFooter: process.env.MAIL_FOOTER ?? "",

  /** Persistent storage for the encrypted Nextcloud app passwords. */
  dataDir: process.env.DATA_DIR ?? "/data",
  /** 32 bytes, Base64 (e.g. `openssl rand -base64 32`). Changing it invalidates all connections. */
  credentialKey: required("CREDENTIAL_KEY"),

  /** From this size (per attachment) CloudShare suggests sending a link. No upper limit – the Nextcloud quota applies. */
  thresholdMb: number("ATTACHMENT_THRESHOLD_MB", 1),
  /**
   * enforce: mails with attachments above the threshold can only be sent via CloudShare (no "Send anyway",
   *          no opt-out per mail).
   * suggest: info bar and a prompt on send that offers "Send anyway".
   */
  attachmentPolicy: oneOf<AttachmentPolicy>("ATTACHMENT_POLICY", ATTACHMENT_POLICIES, "enforce"),
  defaultExpireDays: number("SHARE_EXPIRE_DAYS", 14),
  maxExpireDays: number("SHARE_MAX_EXPIRE_DAYS", 90),

  bwServeUrl: process.env.BW_SERVE_URL ?? "http://127.0.0.1:8087",
  bwPassword: required("BW_PASSWORD"),
  /** Default number of times the password link (Vaultwarden Send) can be opened; the user may change it per mail. */
  sendMaxAccessCount: number("SEND_MAX_ACCESS_COUNT", 3),
  /** Upper limit for the number of accesses a user may choose. */
  sendMaxAccessLimit: number("SEND_MAX_ACCESS_LIMIT", 20),
  sendExpireDays: number("SEND_EXPIRE_DAYS", 7),

  /**
   * Requests per minute and user (oid from the Microsoft token) or, without one, per client address. Generous on
   * purpose – a 5 GB upload is about 500 chunk requests. Expensive actions (Nextcloud login, upload links) get 1/60.
   */
  rateLimitPerMinute: number("RATE_LIMIT_PER_MINUTE", 600),

  /** Nextcloud chunking v2: at least 5 MiB per chunk (except the last), at most 10,000 chunks. */
  chunkSizeBytes: 10 * 1024 * 1024,
};
