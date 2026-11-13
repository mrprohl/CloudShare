import {
  api,
  API_BASE,
  ApiError,
  composeItem,
  getConfig,
  getFileAttachments,
  getToken,
  MB,
  NOT_CONNECTED,
  NOTIFICATION_ID,
  officeAsync,
  readDraftState,
  tagDraft,
  updateDraftState,
  type ServerConfig,
} from "../shared/office";
import { checkPassword, generatePassword } from "../shared/password";
import { errorText } from "../shared/sso";
import { EMPTY_PROFILE, fillFooter, hasPlaceholders, type Profile } from "./footer";
import { filesCard, filesText, requestCard, requestText } from "./mailCard";

// All visible texts are German and follow Outlook's terminology ("Nachricht", "Kennwort").

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const connectView = $<HTMLElement>("connect-view");
const connectButton = $<HTMLButtonElement>("connect");
const connectHint = $<HTMLParagraphElement>("connect-hint");
const connectLink = $<HTMLAnchorElement>("connect-link");
const sendView = $<HTMLElement>("send-view");
const accountText = $<HTMLElement>("account");
const attachmentList = $<HTMLDivElement>("attachments");
const pickedList = $<HTMLDivElement>("picked");
const noAttachments = $<HTMLParagraphElement>("no-attachments");
const fileInput = $<HTMLInputElement>("file");
const passwordInput = $<HTMLInputElement>("password");
const passwordHint = $<HTMLParagraphElement>("password-hint");
const generateButton = $<HTMLButtonElement>("generate");
const revealButton = $<HTMLButtonElement>("reveal");
const modeRandom = $<HTMLInputElement>("mode-random");
const modeOwn = $<HTMLInputElement>("mode-own");
const rulesList = $<HTMLUListElement>("password-rules");
const requestButton = $<HTMLButtonElement>("request");
const tabSend = $<HTMLButtonElement>("tab-send");
const tabShares = $<HTMLButtonElement>("tab-shares");
const panelSend = $<HTMLDivElement>("panel-send");
const panelShares = $<HTMLDivElement>("panel-shares");
const sharesList = $<HTMLDivElement>("shares");
const noShares = $<HTMLParagraphElement>("no-shares");
const refreshSharesButton = $<HTMLButtonElement>("refresh-shares");
const expireInput = $<HTMLInputElement>("expire");
const maxAccessInput = $<HTMLInputElement>("max-access");
const submitButton = $<HTMLButtonElement>("submit");
const optOutBox = $<HTMLInputElement>("opt-out");
const progress = $<HTMLDivElement>("progress");
const progressBar = $<HTMLDivElement>("progress-bar");
const statusBar = $<HTMLDivElement>("status");

let serverConfig: ServerConfig;
/** MAIL_FOOTER with the placeholders filled from the sender's profile (see footer.ts). */
let mailFooter = "";
let footerReady: Promise<void> = Promise.resolve();
/** Files picked from the computer, not yet uploaded. */
let pickedFiles: File[] = [];

// --- Theme ---------------------------------------------------------------------

/** Relative luminance of a #rrggbb colour, 0 (black) … 1 (white). */
function luminance(hex: string): number {
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex.trim());
  if (!m) return 1;
  const [r, g, b] = m.slice(1).map((c) => parseInt(c, 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** Adopts the current Office theme so the pane blends in with Outlook (light/dark, exact background). */
function applyOfficeTheme() {
  const theme = (Office.context as unknown as { officeTheme?: Office.OfficeTheme }).officeTheme;
  const root = document.documentElement;
  const background = theme?.bodyBackgroundColor;
  const dark = background
    ? luminance(background) < 0.5
    : window.matchMedia("(prefers-color-scheme: dark)").matches;
  root.dataset.theme = dark ? "dark" : "light";
  if (background) root.style.setProperty("--bg1", background);
}

// --- Status bar and progress ---------------------------------------------------

type StatusKind = "info" | "success" | "error";

const STATUS_ICONS: Record<StatusKind, string> = {
  info: '<circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor"/><path d="M8 7v4M8 4.8v.4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>',
  success: '<circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor"/><path d="m5.3 8.2 1.8 1.8 3.6-3.8" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>',
  error: '<circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor"/><path d="M8 4.5v4.2M8 11v.4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>',
};

function setStatus(text: string, kind: StatusKind = "info") {
  statusBar.hidden = !text;
  statusBar.className = `messagebar ${kind}`;
  statusBar.innerHTML = text
    ? `<svg class="icon" viewBox="0 0 16 16" aria-hidden="true">${STATUS_ICONS[kind]}</svg><span></span>`
    : "";
  if (text) statusBar.querySelector("span")!.textContent = text;
}

function setProgress(percent: number | null) {
  progress.hidden = percent === null;
  progressBar.style.width = `${percent ?? 0}%`;
}

// --- Helpers -------------------------------------------------------------------

const errorMessage = errorText;

const formatSize = (bytes: number) =>
  bytes >= 1024 ** 3
    ? `${(bytes / 1024 ** 3).toLocaleString("de-DE", { maximumFractionDigits: 1 })} GB`
    : bytes >= MB
      ? `${(bytes / MB).toLocaleString("de-DE", { maximumFractionDigits: 1 })} MB`
      : `${Math.max(1, Math.round(bytes / 1024))} KB`;

const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

// --- Nextcloud connection ------------------------------------------------------

interface ConnectionStatus { connected: boolean; loginName?: string; freeBytes?: number | null }

function showConnection(status: ConnectionStatus) {
  connectView.hidden = status.connected;
  sendView.hidden = !status.connected;
  connectHint.hidden = true;
  const free = typeof status.freeBytes === "number" ? ` · ${formatSize(status.freeBytes)} frei` : "";
  accountText.textContent = `${status.loginName ?? ""}${free}`;
}

async function refreshConnection() {
  showConnection(await api<ConnectionStatus>("/api/nextcloud/status", {}, await getToken()));
}

/** Opens the Nextcloud login in the default browser (usually an SSO session already exists there). */
function openLogin(url: string, popup: Window | null) {
  const ui = Office.context.ui as unknown as { openBrowserWindow?: (url: string) => void };
  if (ui.openBrowserWindow && Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
    ui.openBrowserWindow(url);
  } else if (popup) {
    popup.location.href = url;
  }
}

let connectAttempt = 0;

connectButton.addEventListener("click", async () => {
  const attempt = ++connectAttempt;
  // In the browser (Outlook on the web) open the window immediately, otherwise the popup blocker interferes.
  const usesBrowserApi = Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1");
  const popup = usesBrowserApi ? null : window.open("about:blank", "_blank");

  try {
    setStatus("");
    connectButton.disabled = true;
    const { connectId, loginUrl } = await api<{ connectId: string; loginUrl: string }>(
      "/api/nextcloud/connect",
      { method: "POST" },
      await getToken(),
    );
    openLogin(loginUrl, popup);
    connectLink.href = loginUrl;
    connectHint.hidden = false;

    // Nextcloud discards the login flow after 20 minutes.
    const deadline = Date.now() + 20 * 60_000;
    while (attempt === connectAttempt && Date.now() < deadline) {
      await sleep(2000);
      const result = await api<ConnectionStatus>(
        `/api/nextcloud/connect/${connectId}/poll`,
        { method: "POST" },
        await getToken(),
      );
      if (result.connected) {
        popup?.close();
        await refreshConnection();
        setStatus("Die Verbindung zu Nextcloud wurde hergestellt.", "success");
        return;
      }
    }
    if (attempt === connectAttempt) {
      connectHint.hidden = true;
      setStatus("Die Anmeldung wurde nicht abgeschlossen. Bitte versuchen Sie es erneut.", "error");
    }
  } catch (err) {
    popup?.close();
    connectHint.hidden = true;
    setStatus(`Die Verbindung ist fehlgeschlagen: ${errorMessage(err)}`, "error");
  } finally {
    connectButton.disabled = false;
  }
});

$<HTMLAnchorElement>("disconnect").addEventListener("click", async (e) => {
  e.preventDefault();
  try {
    await api("/api/nextcloud/connection", { method: "DELETE" }, await getToken());
    showConnection({ connected: false });
    setStatus("Die Verbindung zu Nextcloud wurde getrennt.");
  } catch (err) {
    setStatus(`Die Verbindung konnte nicht getrennt werden: ${errorMessage(err)}`, "error");
  }
});

// --- File lists ------------------------------------------------------------------

const FILE_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M5.5 2.5h5.3l3.7 3.7v10.3a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-13a1 1 0 0 1 1-1Z" fill="none" stroke="currentColor"/>' +
  '<path d="M10.5 2.5v4h4" fill="none" stroke="currentColor"/></svg>';

/** One row in Outlook's list style: checkbox, file icon, name, size and an optional "Groß" badge. */
function fileRow(name: string, size: number, value: string, checked: boolean, large: boolean): HTMLLabelElement {
  const row = document.createElement("label");
  row.className = "checkbox list-item";
  row.innerHTML =
    `<input type="checkbox" /><span class="box" aria-hidden="true"></span>${FILE_ICON}` +
    `<span class="file-text"><span class="file-name"></span><span class="file-size"></span></span>` +
    (large ? `<span class="badge">Groß</span>` : "");
  const box = row.querySelector("input")!;
  box.value = value;
  box.checked = checked;
  box.dataset.name = name;
  row.querySelector(".file-name")!.textContent = name;
  row.querySelector<HTMLElement>(".file-name")!.title = name;
  row.querySelector(".file-size")!.textContent = formatSize(size);
  return row;
}

async function renderAttachments() {
  const attachments = await getFileAttachments();
  const threshold = serverConfig.thresholdMb * MB;
  // Large attachments are preselected.
  attachmentList.replaceChildren(
    ...attachments.map((a) => fileRow(a.name, a.size, a.id, a.size > threshold, a.size > threshold)),
  );
  noAttachments.hidden = attachments.length > 0 || pickedFiles.length > 0;
}

function renderPicked() {
  pickedList.replaceChildren(...pickedFiles.map((f, i) => fileRow(f.name, f.size, String(i), true, false)));
  noAttachments.hidden = attachmentList.childElementCount > 0 || pickedFiles.length > 0;
}

$<HTMLButtonElement>("add-files").addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  pickedFiles = [...pickedFiles, ...Array.from(fileInput.files ?? [])];
  fileInput.value = "";
  renderPicked();
});

function base64ToBlob(base64: string): Blob {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes]);
}

async function readAttachment(id: string): Promise<Blob> {
  const content = await officeAsync<Office.AttachmentContent>((cb) =>
    composeItem().getAttachmentContentAsync(id, cb),
  );
  if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
    throw new Error("Dieser Anhangstyp kann nicht übertragen werden.");
  }
  return base64ToBlob(content.content);
}

// --- Upload and share ---------------------------------------------------------

interface UploadStart { uploadId: string; chunkSize: number; totalChunks: number }
interface UploadResult { shareUrl: string; expireDate: string; jobId: string; fileName: string; size: number }

async function upload(blob: Blob, fileName: string): Promise<UploadResult> {
  const start = await api<UploadStart>(
    "/api/uploads",
    { method: "POST", body: JSON.stringify({ fileName, size: blob.size }) },
    await getToken(),
  );

  try {
    for (let index = 1; index <= start.totalChunks; index++) {
      const chunk = blob.slice((index - 1) * start.chunkSize, index * start.chunkSize);
      // Fetch the token per chunk: Office caches it and renews it during long uploads.
      await api(
        `/api/uploads/${start.uploadId}/chunks/${index}`,
        { method: "PUT", body: chunk, headers: { "Content-Type": "application/octet-stream" } },
        await getToken(),
      );
      const percent = Math.round((index / start.totalChunks) * 100);
      setProgress(percent);
      setStatus(`„${fileName}“ wird hochgeladen … ${percent} %`);
    }
  } catch (err) {
    await api(`/api/uploads/${start.uploadId}`, { method: "DELETE" }, await getToken()).catch(() => undefined);
    throw err;
  }

  const state = await readDraftState();
  const result = await api<UploadResult>(
    `/api/uploads/${start.uploadId}/complete`,
    {
      method: "POST",
      body: JSON.stringify({
        password: passwordInput.value,
        expireDays: Number(expireInput.value),
        maxAccess: maxAccessValue(),
        jobId: state.jobId,
      }),
    },
    await getToken(),
  );
  // From the first file on, the password applies to the whole mail.
  await updateDraftState({ jobId: result.jobId, password: passwordInput.value, maxAccess: maxAccessValue() });
  await tagDraft(result.jobId);
  lockPassword();
  return result;
}

const formatDate = (isoDate: string) => new Date(isoDate).toLocaleDateString("de-DE");

/**
 * Fills MAIL_FOOTER from the sender's Microsoft 365 profile. If the profile cannot be read, name and address
 * known to Outlook are used, so the footer never shows raw placeholders.
 */
async function resolveFooter() {
  const template = serverConfig.mailFooter;
  if (!hasPlaceholders(template)) {
    mailFooter = template;
    return;
  }
  let profile: Profile;
  try {
    profile = await api<Profile>("/api/profile", {}, await getToken());
  } catch {
    const user = Office.context.mailbox.userProfile;
    profile = { ...EMPTY_PROFILE, displayName: user.displayName, email: user.emailAddress };
  }
  mailFooter = fillFooter(template, profile);
}

/** Inserts a card (HTML mail) or plain text (text mail) at the cursor. */
async function insertIntoBody(html: string, text: string) {
  const body = composeItem().body;
  const type = await officeAsync<Office.CoercionType>((cb) => body.getTypeAsync(cb)).catch(
    () => Office.CoercionType.Html,
  );
  const isText = type === Office.CoercionType.Text;
  await officeAsync<void>((cb) =>
    body.setSelectedDataAsync(isText ? text : html, { coercionType: isText ? Office.CoercionType.Text : Office.CoercionType.Html }, cb),
  );
}

/** Inserts the shared files as a card with download buttons (see mailCard.ts). */
async function insertLinks(results: UploadResult[]) {
  const files = results.map((r) => ({ name: r.fileName, size: r.size, url: r.shareUrl }));
  const expire = formatDate(results[0].expireDate);
  await footerReady;
  await insertIntoBody(
    filesCard(files, expire, mailFooter, formatSize),
    filesText(files, expire, mailFooter, formatSize),
  );
}

// --- Password: random (recommended) or own ---------------------------------------

function setRevealed(revealed: boolean) {
  passwordInput.type = revealed ? "text" : "password";
  revealButton.setAttribute("aria-pressed", String(revealed));
  const label = revealed ? "Kennwort ausblenden" : "Kennwort anzeigen";
  revealButton.title = label;
  revealButton.setAttribute("aria-label", label);
}

/** Live feedback for an own password; rules are only marked as failed once something has been typed. */
function renderRules() {
  const check = checkPassword(passwordInput.value);
  const typed = passwordInput.value.length > 0;
  rulesList.querySelectorAll<HTMLLIElement>("li").forEach((li) => {
    const ok = check[li.dataset.rule as "length" | "classes" | "variety"];
    li.classList.toggle("ok", ok);
    li.classList.toggle("fail", typed && !ok);
  });
}

/** Random: generated, visible, read-only. Own: hidden input with "show" toggle and live rule check. */
function setPasswordMode(mode: "random" | "own") {
  const own = mode === "own";
  modeRandom.checked = !own;
  modeOwn.checked = own;
  passwordInput.readOnly = !own;
  generateButton.hidden = own;
  revealButton.hidden = !own;
  rulesList.hidden = !own;
  if (own) {
    passwordInput.value = "";
    setRevealed(false);
    renderRules();
    passwordInput.focus();
  } else {
    passwordInput.type = "text";
    passwordInput.value = generatePassword();
  }
}

modeRandom.addEventListener("change", () => setPasswordMode("random"));
modeOwn.addEventListener("change", () => setPasswordMode("own"));
passwordInput.addEventListener("input", renderRules);
revealButton.addEventListener("click", () => setRevealed(passwordInput.type === "password"));

/** From the first upload on, the password applies to the whole mail and can no longer be changed. */
function lockPassword() {
  passwordInput.readOnly = true;
  modeRandom.disabled = true;
  modeOwn.disabled = true;
  generateButton.hidden = true;
  rulesList.hidden = true;
  revealButton.hidden = false;
  setRevealed(false);
  // One Vaultwarden Send per mail – its number of accesses is fixed together with the password.
  maxAccessInput.readOnly = true;
  passwordHint.textContent =
    "Dieses Kennwort gilt für alle Dateien dieser Nachricht. Die Empfänger erhalten es nach dem Senden automatisch.";
}

/** Number of accesses within the allowed range (the server limits it again). */
function maxAccessValue(): number {
  const value = Math.floor(Number(maxAccessInput.value));
  if (!Number.isFinite(value) || value < 1) return serverConfig.sendMaxAccessCount;
  return Math.min(value, serverConfig.sendMaxAccessLimit);
}

submitButton.addEventListener("click", async () => {
  const selectedAttachments = Array.from(attachmentList.querySelectorAll<HTMLInputElement>("input:checked"));
  const selectedFiles = Array.from(pickedList.querySelectorAll<HTMLInputElement>("input:checked")).map(
    (box) => pickedFiles[Number(box.value)],
  );
  if (selectedAttachments.length + selectedFiles.length === 0) {
    return setStatus("Wählen Sie mindestens eine Datei aus.", "error");
  }
  if (!passwordInput.readOnly && !checkPassword(passwordInput.value).ok) {
    renderRules();
    passwordInput.focus();
    return setStatus("Das Kennwort erfüllt noch nicht alle Anforderungen.", "error");
  }

  submitButton.disabled = true;
  const results: UploadResult[] = [];
  let failed = false;
  try {
    for (const box of selectedAttachments) {
      const name = box.dataset.name ?? "Anhang";
      setProgress(0);
      setStatus(`„${name}“ wird vorbereitet …`);
      const blob = await readAttachment(box.value);
      results.push(await upload(blob, name));
      // Remove from the mail only after a successful upload.
      await officeAsync<void>((cb) => composeItem().removeAttachmentAsync(box.value, cb));
    }
    for (const file of selectedFiles) {
      if (file.size === 0) throw new Error(`„${file.name}“ ist leer.`);
      setProgress(0);
      results.push(await upload(file, file.name));
      pickedFiles = pickedFiles.filter((f) => f !== file);
    }
  } catch (err) {
    failed = true;
    if (err instanceof ApiError && err.code === NOT_CONNECTED) showConnection({ connected: false });
    setStatus(errorMessage(err), "error");
  } finally {
    // Link already uploaded files anyway.
    if (results.length > 0) {
      try {
        await insertLinks(results);
        if (!failed) {
          setStatus(
            `${results.length === 1 ? "Der Link wurde" : `${results.length} Links wurden`} eingefügt. ` +
              `Das Kennwort wird nach dem Senden automatisch verschickt.`,
            "success",
          );
        }
      } catch (err) {
        setStatus(`Der Link konnte nicht eingefügt werden: ${errorMessage(err)}`, "error");
      }
      renderPicked();
      await renderAttachments().catch(() => undefined);
      await refreshConnection().catch(() => undefined); // refreshes the free quota
    }
    submitButton.disabled = false;
    setProgress(null);
  }
});

generateButton.addEventListener("click", () => {
  passwordInput.value = generatePassword();
});

// --- Upload link ("Dateien anfordern") ----------------------------------------------

requestButton.addEventListener("click", async () => {
  if (!passwordInput.readOnly && !checkPassword(passwordInput.value).ok) {
    renderRules();
    passwordInput.focus();
    return setStatus("Das Kennwort erfüllt noch nicht alle Anforderungen.", "error");
  }
  requestButton.disabled = true;
  try {
    setStatus("Upload-Link wird erstellt …");
    const state = await readDraftState();
    const result = await api<{ shareUrl: string; expireDate: string; jobId: string }>(
      "/api/requests",
      {
        method: "POST",
        body: JSON.stringify({
          password: passwordInput.value,
          expireDays: Number(expireInput.value),
          maxAccess: maxAccessValue(),
          jobId: state.jobId,
        }),
      },
      await getToken(),
    );
    // Same password and password mail as the files of this mail.
    await updateDraftState({ jobId: result.jobId, password: passwordInput.value, maxAccess: maxAccessValue() });
    await tagDraft(result.jobId);
    lockPassword();
    const expire = formatDate(result.expireDate);
    await footerReady;
    await insertIntoBody(
      requestCard(result.shareUrl, expire, mailFooter),
      requestText(result.shareUrl, expire, mailFooter),
    );
    setStatus("Der Upload-Link wurde eingefügt. Das Kennwort wird nach dem Senden automatisch verschickt.", "success");
  } catch (err) {
    if (err instanceof ApiError && err.code === NOT_CONNECTED) showConnection({ connected: false });
    setStatus(errorMessage(err), "error");
  } finally {
    requestButton.disabled = false;
  }
});

// --- Tabs ------------------------------------------------------------------------------

function selectTab(tab: "send" | "shares") {
  const shares = tab === "shares";
  tabSend.setAttribute("aria-selected", String(!shares));
  tabShares.setAttribute("aria-selected", String(shares));
  panelSend.hidden = shares;
  panelShares.hidden = !shares;
  if (shares) loadShares().catch((err) => setStatus(errorMessage(err), "error"));
}

tabSend.addEventListener("click", () => selectTab("send"));
tabShares.addEventListener("click", () => selectTab("shares"));

// --- My shares ("Meine Freigaben") -----------------------------------------------------

interface ShareItem {
  id: string;
  kind: "file" | "request";
  name: string;
  url: string;
  sharedAt: number;
  expiration: string | null;
  downloads: { count: number; last: string | null } | null;
  received: number | null;
}

const REQUEST_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M10 13V4m0 0L6.5 7.5M10 4l3.5 3.5M4 12v3a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-3" fill="none" ' +
  'stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';

const formatDateTime = (iso: string) =>
  new Date(iso).toLocaleString("de-DE", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });

/** Status line: downloads for files, received uploads for upload links, or "abgelaufen". */
function shareState(share: ShareItem): { text: string; done: boolean; expired: boolean } {
  const expired = share.expiration !== null && share.expiration < new Date().toISOString().slice(0, 10);
  if (expired) return { text: "Abgelaufen", done: false, expired: true };
  if (share.kind === "request") {
    if (share.received === null) return { text: "Status unbekannt", done: false, expired: false };
    return share.received > 0
      ? { text: `${share.received} ${share.received === 1 ? "Datei" : "Dateien"} empfangen`, done: true, expired: false }
      : { text: "Noch keine Dateien empfangen", done: false, expired: false };
  }
  if (!share.downloads) return { text: "Download-Status nicht verfügbar", done: false, expired: false };
  return share.downloads.count > 0
    ? {
        text:
          `${share.downloads.count}× heruntergeladen` +
          (share.downloads.last ? `, zuletzt ${formatDateTime(share.downloads.last)}` : ""),
        done: true,
        expired: false,
      }
    : { text: "Noch nicht heruntergeladen", done: false, expired: false };
}

/** Copies text; the async clipboard API is not available in every Outlook version. */
async function copyText(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
}

function shareRow(share: ShareItem): HTMLElement {
  const row = document.createElement("div");
  row.className = "share";
  const state = shareState(share);
  row.innerHTML =
    (share.kind === "request" ? REQUEST_ICON : FILE_ICON) +
    `<div class="share-body">` +
    `<div class="share-name"></div>` +
    `<div class="share-meta"></div>` +
    `<div class="share-state${state.done ? " done" : ""}${state.expired ? " expired" : ""}">` +
    `<span class="dot-state" aria-hidden="true"></span><span class="state-text"></span></div>` +
    `<div class="share-actions">` +
    `<button type="button" class="btn btn-subtle copy">Link kopieren</button>` +
    `<button type="button" class="btn btn-subtle btn-danger revoke">Zurückziehen</button>` +
    `</div></div>`;
  row.querySelector<HTMLElement>(".share-name")!.textContent = share.name;
  row.querySelector<HTMLElement>(".share-name")!.title = share.name;
  row.querySelector<HTMLElement>(".share-meta")!.textContent =
    `Freigegeben am ${new Date(share.sharedAt).toLocaleDateString("de-DE")}` +
    (share.expiration ? ` · gültig bis ${formatDate(share.expiration)}` : "");
  row.querySelector<HTMLElement>(".state-text")!.textContent = state.text;

  const copy = row.querySelector<HTMLButtonElement>(".copy")!;
  copy.addEventListener("click", async () => {
    await copyText(share.url);
    copy.textContent = "Kopiert";
    setTimeout(() => (copy.textContent = "Link kopieren"), 2000);
  });

  // Two-step confirmation instead of a modal dialog (add-ins cannot show native dialogs).
  const revoke = row.querySelector<HTMLButtonElement>(".revoke")!;
  let armedUntil = 0;
  revoke.addEventListener("click", async () => {
    if (Date.now() > armedUntil) {
      armedUntil = Date.now() + 4000;
      revoke.textContent = "Wirklich zurückziehen?";
      revoke.classList.add("confirm");
      setTimeout(() => {
        if (Date.now() >= armedUntil) {
          revoke.textContent = "Zurückziehen";
          revoke.classList.remove("confirm");
        }
      }, 4100);
      return;
    }
    revoke.disabled = true;
    try {
      await api(`/api/shares/${encodeURIComponent(share.id)}`, { method: "DELETE" }, await getToken());
      row.remove();
      setStatus(`„${share.name}“ ist nicht mehr abrufbar. Die Datei bleibt in Ihrer Nextcloud.`, "success");
      noShares.hidden = sharesList.childElementCount > 0;
    } catch (err) {
      revoke.disabled = false;
      setStatus(`Die Freigabe konnte nicht zurückgezogen werden: ${errorMessage(err)}`, "error");
    }
  });
  return row;
}

async function loadShares() {
  refreshSharesButton.disabled = true;
  try {
    const shares = await api<ShareItem[]>("/api/shares", {}, await getToken());
    sharesList.replaceChildren(...shares.map(shareRow));
    noShares.hidden = shares.length > 0;
  } finally {
    refreshSharesButton.disabled = false;
  }
}

refreshSharesButton.addEventListener("click", () => {
  loadShares().catch((err) => setStatus(errorMessage(err), "error"));
});

optOutBox.addEventListener("change", async () => {
  await updateDraftState({ optOut: optOutBox.checked });
  if (optOutBox.checked) {
    // Remove the info bar; it reappears only when the box is unchecked and attachments change.
    await officeAsync<void>((cb) => composeItem().notificationMessages.removeAsync(NOTIFICATION_ID, cb)).catch(
      () => undefined,
    );
    setStatus("Diese Nachricht wird ohne CloudShare gesendet.");
  } else {
    setStatus("");
  }
});

Office.onReady(async () => {
  applyOfficeTheme();
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyOfficeTheme);
  } catch {
    // Theme change events are not available in every Outlook version.
  }

  try {
    serverConfig = await getConfig();
  } catch {
    setStatus(`CloudShare ist derzeit nicht erreichbar (${API_BASE}).`, "error");
    return;
  }
  footerReady = resolveFooter().catch(() => undefined); // in the background – needed when a link is inserted

  const state = await readDraftState();
  maxAccessInput.max = String(serverConfig.sendMaxAccessLimit);
  maxAccessInput.value = String(state.maxAccess ?? serverConfig.sendMaxAccessCount);
  if (state.password) {
    passwordInput.value = state.password;
    lockPassword();
  } else {
    setPasswordMode("random");
  }
  optOutBox.checked = Boolean(state.optOut);
  // Under "enforce" there is no way to send large attachments without CloudShare.
  $<HTMLLabelElement>("opt-out-row").hidden = serverConfig.attachmentPolicy === "enforce";
  expireInput.value = String(serverConfig.defaultExpireDays);
  expireInput.max = String(serverConfig.maxExpireDays);

  await renderAttachments().catch(() => undefined);
  composeItem().addHandlerAsync(Office.EventType.AttachmentsChanged, () => {
    renderAttachments().catch(() => undefined);
  });

  try {
    await refreshConnection();
  } catch (err) {
    setStatus(`Die Anmeldung ist fehlgeschlagen: ${errorMessage(err)}`, "error");
  }
});
