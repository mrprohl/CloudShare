import { toSsoError } from "./sso";

/** Replaced with the real base URL at container start (see docker/entrypoint.sh). */
export const API_BASE = "%%PUBLIC_BASE_URL%%";

export interface ServerConfig {
  thresholdMb: number;
  /** enforce: large attachments only via CloudShare; suggest: prompt with "Send anyway". */
  attachmentPolicy: "enforce" | "suggest";
  defaultExpireDays: number;
  maxExpireDays: number;
  chunkSize: number;
  /** Default and upper limit for how often the password link (Vaultwarden Send) can be opened. */
  sendMaxAccessCount: number;
  sendMaxAccessLimit: number;
  /** Optional footer line of the card inserted into the mail. */
  mailFooter: string;
  /** Diagnostics switch (CLIENT_TRACE): the send handler reports its steps to the server. */
  trace?: boolean;
}

/** Office SSO: "a previous getAccessToken call is still in progress". */
const SSO_IN_PROGRESS = 13008;

let cachedToken: { token: string; expiresAt: number } | null = null;
let pendingToken: Promise<string> | null = null;

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

/** Rejects with a German message naming `step` if `promise` does not settle within `ms`. */
export function withTimeout<T>(promise: Promise<T>, ms: number, step: string): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const timer = setTimeout(
      () => reject(new Error(`Zeitüberschreitung: ${step} hat nach ${Math.round(ms / 1000)} s nicht geantwortet.`)),
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
        forMSGraphAccess: true,
      });
      cachedToken = { token, expiresAt: tokenExpiry(token) };
      return token;
    } catch (err) {
      // Another runtime of this add-in (taskpane / send handler) may be asking at the same moment.
      const code = err && typeof err === "object" ? Number((err as { code?: unknown }).code) : NaN;
      if (code === SSO_IN_PROGRESS && attempt < 4) {
        await delay(400 * attempt);
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
 * German message and the code (see sso.ts).
 */
export function getToken(): Promise<string> {
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

/** Tests only: forget the cached token. */
export function resetTokenCache(): void {
  cachedToken = null;
  pendingToken = null;
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
  if (token) headers.Authorization = `Bearer ${token}`;
  if (init.body && typeof init.body === "string") headers["Content-Type"] = "application/json";
  const res = await fetch(API_BASE + path, { ...init, headers });
  const json = (await res.json().catch(() => ({}))) as { error?: string; code?: string };
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
    const config = await withTimeout(getConfig(), TIMEOUTS.config, "Der CloudShare-Server");
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
  password?: string;
  /** How often the password link can be opened – changeable until sending (PATCH /api/jobs/:id). */
  maxAccess?: number;
  /** Links of this mail were shared "Ohne Kennwort" – no password job, no password mail. */
  noPassword?: boolean;
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
  const fresh: DraftState = {};
  await officeAsync<void>((cb) => composeItem().sessionData.setAsync(STATE_KEY, JSON.stringify(fresh), cb));
  return true;
}
