import { language, t } from "./i18n";
import { toSsoError } from "./sso";

/** Replaced with the real base URL at container start (see docker/entrypoint.sh). */
export const API_BASE = "%%PUBLIC_BASE_URL%%";

export interface ServerConfig {
  thresholdMb: number;
  /** enforce: large attachments only via CloudShare; suggest: prompt with "Send anyway". */
  attachmentPolicy: "enforce" | "suggest";
  defaultExpireDays: number;
  maxExpireDays: number;
  /** Longest validity of a link with password (as long as its password in Vaultwarden, at most 30); older servers: none. */
  passwordLinkMaxDays?: number;
  chunkSize: number;
  /** Default and upper limit for how often the password link (Vaultwarden Send) can be opened. */
  sendMaxAccessCount: number;
  sendMaxAccessLimit: number;
  /** Diagnostics switch (CLIENT_TRACE): the send handler reports its steps to the server. */
  trace?: boolean;
  /** SMS channel configured on the server (seven.io); country code for numbers entered without one. */
  smsEnabled?: boolean;
  smsDefaultCountry?: string;
  /** SMS_ALLOWED_COUNTRIES: calling codes SMS may go to and their country names as text ([] / "" = all countries). */
  smsCountryCodes?: string[];
  smsCountries?: string;
  /** SMS_ALLOWED_NUMBERS / SMS_BLOCKED_NUMBERS: regular expressions on the E.164 number ("" = none). */
  smsAllowedNumbers?: string;
  smsBlockedNumbers?: string;
  /** Confidential message: longest text (characters). */
  messageMaxLength?: number;
  /** Password link and confidential message (Vaultwarden Sends): default and maximum lifetime in days. */
  sendExpireDays?: number;
  sendMaxExpireDays?: number;
  /** Messages disappear after this many seconds (0: they stay; progress always stays); default 5. */
  successMessageSeconds?: number;
  /** NEXTCLOUD_URL – "Nextcloud öffnen" in the account menu. */
  nextcloudUrl?: string;
}

export type Channel = "recipients" | "mail" | "sms";

/** Personal settings (⚙, server/src/settings.ts). */
export interface UserSettings {
  downloadNotice: boolean;
  uploadNotice: boolean;
  smsNotice: boolean;
  /** Ways "Kennwort senden an" offers, and the one preselected for a new mail. */
  channels: Channel[];
  defaultChannel: Channel;
  /** "auto" (as in Outlook) or a language code. */
  language: string;
  /** Preselected language for the recipients (link card, password mail, SMS) – changeable per mail. */
  recipientLanguage: string;
  /** Own signature of the password mails (plain text, empty line = new paragraph); null = the organisation's (MAIL_FOOTER). */
  signature: { format: "text"; content: string } | null;
}

/** GET/PUT/DELETE /api/settings: what applies, the organisation's defaults and what the server offers at all. */
export interface SettingsResponse {
  settings: UserSettings;
  defaults: UserSettings;
  available: { downloadNotice: boolean; uploadNotice: boolean; smsNotice: boolean; channels: Channel[]; languages: string[] };
}

/** Office SSO: "a previous getAccessToken call is still in progress". */
const SSO_IN_PROGRESS = 13008;
/** Office SSO: "too many calls in a short time" – Outlook throttles getAccessToken per add-in. */
const SSO_THROTTLED = 13013;

type CachedToken = { token: string; expiresAt: number };
let cachedToken: CachedToken | null = null;
let pendingToken: Promise<string> | null = null;

/**
 * The token also survives a reload of the pane (classic Outlook loads it anew for every mail selected while it is not
 * pinned; a language change reloads it): kept in sessionStorage – this Outlook window only, gone when it closes – so
 * clicking through mails does not ask Office every time and run into its limit (13013). Best effort: the JavaScript-
 * only runtime of classic Outlook has no storage.
 */
const TOKEN_KEY = "cloudshare-sso";
function storedToken(): CachedToken | null {
  try {
    const stored = JSON.parse(sessionStorage.getItem(TOKEN_KEY) || "null") as CachedToken | null;
    return stored && typeof stored.token === "string" && typeof stored.expiresAt === "number" ? stored : null;
  } catch {
    return null;
  }
}
function storeToken(value: CachedToken | null) {
  try {
    if (value) sessionStorage.setItem(TOKEN_KEY, JSON.stringify(value));
    else sessionStorage.removeItem(TOKEN_KEY);
  } catch {
    // no sessionStorage here
  }
}

/** Expiry of a JWT in ms; unreadable tokens are reused for five minutes at most. */
function tokenExpiry(token: string): number {
  try {
    const payload = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
    const exp = Number(JSON.parse(atob(payload)).exp);
    if (Number.isFinite(exp)) return exp * 1000;
  } catch {
    // atob may be missing in the JavaScript-only runtime of classic Outlook
  }
  return Date.now() + 5 * 60_000;
}

const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

/**
 * Time limits of the send handler (ms). Outlook only shows "takes longer than expected" when a step never
 * answers; with a limit the user gets a message naming the step instead. Changeable for tests.
 */
export const TIMEOUTS = { config: 5_000, token: 15_000, server: 20_000, office: 5_000 };

/** Rejects with a message naming `step` if `promise` does not settle within `ms`. */
export function withTimeout<T>(promise: Promise<T>, ms: number, step: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error(t("timeout", { step, seconds: Math.round(ms / 1000) }))),
      ms,
    );
    promise.then(
      (value) => {
        clearTimeout(timer);
        resolve(value);
      },
      (err) => {
        clearTimeout(timer);
        reject(err);
      },
    );
  });
}

async function requestToken(): Promise<string> {
  for (let attempt = 1; ; attempt++) {
    try {
      const token = await Office.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        // No forMSGraphAccess: with it Office refuses the token (13012) whenever it does not see the admin consent
        // for the Graph scopes in its own deployment record – even for a correctly deployed add-in whose consent
        // was granted in Entra. Missing consent shows up instead on the server, when the On-Behalf-Of exchange for
        // Graph fails (HttpError 403 with the Entra message, see onBehalfOf in server/src/auth.ts).
      });
      cachedToken = { token, expiresAt: tokenExpiry(token) };
      storeToken(cachedToken);
      return token;
    } catch (err) {
      // Another runtime of this add-in (taskpane / send handler) may be asking at the same moment.
      const code = err && typeof err === "object" ? Number((err as { code?: unknown }).code) : NaN;
      if (code === SSO_IN_PROGRESS && attempt < 4) {
        await delay(400 * attempt);
        continue;
      }
      // Too many requests in a short time: Outlook lets the add-in ask again after a few seconds.
      if (code === SSO_THROTTLED && attempt < 3) {
        await delay(3000 * attempt);
        continue;
      }
      throw toSsoError(err);
    }
  }
}

/**
 * SSO token for the backend (audience = our own app registration). Office allows only one getAccessToken call at
 * a time (otherwise error 13008), so concurrent callers share one request and a valid token is reused until shortly
 * before it expires. Office rejects with a plain object ({ code, message }); it is turned into an Error with a
 * readable message and the code (see sso.ts).
 */
export function getToken(): Promise<string> {
  if (!cachedToken) cachedToken = storedToken(); // from before a reload of the pane
  if (cachedToken && cachedToken.expiresAt - 2 * 60_000 > Date.now()) return Promise.resolve(cachedToken.token);
  // No Promise.finally / Headers: the JavaScript-only runtime of classic Outlook is not a full browser.
  if (!pendingToken) {
    pendingToken = requestToken().then(
      (token) => {
        pendingToken = null;
        return token;
      },
      (err) => {
        pendingToken = null;
        throw err;
      },
    );
  }
  return pendingToken;
}

/** Tests only: forget the cached token – `keepStored`: as a reload of the pane does (sessionStorage stays). */
export function resetTokenCache(keepStored = false): void {
  cachedToken = null;
  pendingToken = null;
  if (!keepStored) storeToken(null);
}

export const NOT_CONNECTED = "NEXTCLOUD_NOT_CONNECTED";

export class ApiError extends Error {
  constructor(
    message: string,
    public readonly code?: string,
  ) {
    super(message);
  }
}

export async function api<T>(path: string, init: RequestInit = {}, token?: string): Promise<T> {
  const headers: Record<string, string> = { ...(init.headers as Record<string, string> | undefined) };
  headers["X-CloudShare-Language"] = language(); // error messages in the add-in's language (server/src/i18n.ts)
  if (token) headers.Authorization = `Bearer ${token}`;
  if (init.body && typeof init.body === "string") headers["Content-Type"] = "application/json";
  const res = await fetch(API_BASE + path, { ...init, headers });
  const json = (await res.json().catch(() => ({}))) as { error?: string; code?: string };
  // The server refused the sign-in (not Nextcloud): the kept token is no good – the next call asks Office again.
  if (res.status === 401 && token && json.code !== NOT_CONNECTED && cachedToken?.token === token) {
    cachedToken = null;
    storeToken(null);
  }
  if (!res.ok) throw new ApiError(json.error || `HTTP ${res.status}`, json.code);
  return json as T;
}

export const getConfig = () => api<ServerConfig>("/api/config");

export const CONFIG_CACHE_KEY = "cloudshare-config";

/**
 * Server configuration for the event handlers. The last known value is kept in roamingSettings so that the
 * attachment policy is still applied while the server is unreachable.
 */
export async function getConfigCached(): Promise<ServerConfig> {
  try {
    const config = await withTimeout(getConfig(), TIMEOUTS.config, t("step.server"));
    try {
      Office.context.roamingSettings.set(CONFIG_CACHE_KEY, config);
      Office.context.roamingSettings.saveAsync();
    } catch {
      // Caching is best effort.
    }
    return config;
  } catch (err) {
    const cached = Office.context.roamingSettings.get(CONFIG_CACHE_KEY) as ServerConfig | undefined;
    if (cached) return cached;
    throw err;
  }
}

/** Promise wrapper around the callback-based Office.js APIs. */
export function officeAsync<T>(call: (cb: (r: Office.AsyncResult<T>) => void) => void): Promise<T> {
  return new Promise((resolve, reject) =>
    call((r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error))),
  );
}

export const MB = 1024 * 1024;

/** ID of the taskpane button in the manifest – target of the notification and Smart Alert buttons. */
export const TASKPANE_COMMAND_ID = "cloudshareOpenPane";
export const NOTIFICATION_ID = "cloudshare-large-attachment";

export const composeItem = () => Office.context.mailbox.item as Office.MessageCompose;

/** Real file attachments (no inline images, no attached mails/cloud links). */
export async function getFileAttachments(): Promise<Office.AttachmentDetailsCompose[]> {
  const all = await officeAsync<Office.AttachmentDetailsCompose[]>((cb) => composeItem().getAttachmentsAsync(cb));
  return all.filter((a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline);
}

/**
 * Name of the custom property that links a draft – and its copy in "Sent Items" – to its password job.
 * The server looks for it to send the password mail only after the mail has really been sent.
 * Must match JOB_PROPERTY in server/src/graph.ts.
 */
export const JOB_PROPERTY = "cloudshareJob";

/** Tags the draft with the job ID (stored with the item, travels into "Sent Items"). */
export async function tagDraft(jobId: string): Promise<void> {
  const props = await officeAsync<Office.CustomProperties>((cb) => composeItem().loadCustomPropertiesAsync(cb));
  if (props.get(JOB_PROPERTY) === jobId) return;
  props.set(JOB_PROPERTY, jobId);
  await officeAsync<void>((cb) => props.saveAsync(cb));
}

// Per-draft state in sessionData (Mailbox 1.11): pending password job, its password
// (all files of one mail share password and password mail), the inserted links and the "Ohne Kennwort" choice.
const STATE_KEY = "cloudshare-state";

export interface DraftState {
  jobId?: string;
  /** "Sprache für die Empfänger" chosen for this mail (else the one from the settings). */
  recipientLanguage?: string;
  password?: string;
  /** How often the password link can be opened – changeable until sending (PATCH /api/jobs/:id). */
  maxAccess?: number;
  /** Links of this mail were shared "Ohne Kennwort" – no password job, no password mail. */
  noPassword?: boolean;
  /** "Kennwort an andere Adresse": the password mail goes only to this address (set on the job by the taskpane). */
  passwordTo?: string;
  /** "Kennwort per SMS": the password goes by SMS to this number (E.164) instead of a password mail. */
  passwordSms?: string;
  /** Mail with one password per "Als Link einfügen" step (every step has its own password and password mail). */
  separate?: boolean;
  /**
   * The last step with a password: "Gleiches Kennwort wie zuvor" (preselected after it) puts the next link into this
   * step – same password, same password mail or SMS. Only in the draft's session data, like `password` always was.
   */
  lastStep?: { id: string; password: string };
  /**
   * Share tokens of the links CloudShare inserted (last URL segment). If none of them is left in the body, the user
   * deleted the links: password, job and tag are released. Drafts from before this field are never released.
   */
  links?: string[];
}

export async function readDraftState(): Promise<DraftState> {
  try {
    const raw = await officeAsync<string>((cb) => composeItem().sessionData.getAsync(STATE_KEY, cb));
    return JSON.parse(raw) as DraftState;
  } catch {
    return {};
  }
}

export async function updateDraftState(patch: DraftState): Promise<DraftState> {
  const next = { ...(await readDraftState()), ...patch };
  await officeAsync<void>((cb) => composeItem().sessionData.setAsync(STATE_KEY, JSON.stringify(next), cb));
  return next;
}

/** Remembers a link CloudShare is about to insert (see DraftState.links). */
export async function rememberLink(token: string): Promise<void> {
  const state = await readDraftState();
  const links = state.links ?? [];
  if (!links.includes(token)) await updateDraftState({ links: [...links, token] });
}

/**
 * All CloudShare links deleted from the mail? Then the draft starts over: no password, no job, no tag – so neither a
 * password mail goes out nor is the password option locked. Returns true if the draft was released.
 * Unreadable body or drafts without recorded links: nothing is released (a password mail must never get lost).
 */
export async function releaseDraftIfLinksGone(): Promise<boolean> {
  const state = await readDraftState();
  if (!state.links?.length || (!state.jobId && !state.noPassword)) return false;
  let body: string;
  try {
    body = await officeAsync<string>((cb) => composeItem().body.getAsync(Office.CoercionType.Html, cb));
  } catch {
    return false;
  }
  if (state.links.some((token) => body.includes(token))) return false;
  if (state.jobId) {
    const props = await officeAsync<Office.CustomProperties>((cb) => composeItem().loadCustomPropertiesAsync(cb));
    props.remove(JOB_PROPERTY);
    await officeAsync<void>((cb) => props.saveAsync(cb));
  }
  // The other password recipient is a setting of the mail, not of the deleted links – keep it.
  const fresh: DraftState = state.passwordTo ? { passwordTo: state.passwordTo } : {};
  await officeAsync<void>((cb) => composeItem().sessionData.setAsync(STATE_KEY, JSON.stringify(fresh), cb));
  return true;
}
