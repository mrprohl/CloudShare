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
const modeNone = $<HTMLInputElement>("mode-none");
const passwordGroup = $<HTMLDivElement>("password-group");
const maxAccessField = $<HTMLDivElement>("max-access-field");
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

interface ConnectionStatus {
  connected: boolean;
  loginName?: string;
  /** Nextcloud storage in bytes; null = unlimited or unknown. */
  freeBytes?: number | null;
  usedBytes?: number | null;
  totalBytes?: number | null;
}

/** Storage behind the account: "[12,3 GB von 50 GB frei]", "[12,3 GB frei]" (no quota) or "[3,2 GB belegt]". */
function storageText(status: ConnectionStatus): string {
  const isBytes = (value: unknown): value is number => typeof value === "number" && value >= 0;
  if (isBytes(status.freeBytes) && isBytes(status.totalBytes)) {
    return ` [${formatSize(status.freeBytes)} von ${formatSize(status.totalBytes)} frei]`;
  }
  if (isBytes(status.freeBytes)) return ` [${formatSize(status.freeBytes)} frei]`;
  if (isBytes(status.usedBytes)) return ` [${formatSize(status.usedBytes)} belegt]`;
  return "";
}

function showConnection(status: ConnectionStatus) {
  connectView.hidden = status.connected;
  sendView.hidden = !status.connected;
  connectHint.hidden = true;
  accountText.textContent = `${status.loginName ?? ""}${storageText(status)}`;
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
/** jobId is null for links "Ohne Kennwort" – no password mail. */
interface UploadResult { shareUrl: string; expireDate: string; jobId: string | null; fileName: string; size: number }

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

  const result = await api<UploadResult>(
    `/api/uploads/${start.uploadId}/complete`,
    { method: "POST", body: JSON.stringify(await shareSettingsBody()) },
    await getToken(),
  );
  await rememberShareChoice(result.jobId);
  return result;
}

/** Password settings of the "Senden" tab for a new link of this mail (upload, upload link, link again). */
async function shareSettingsBody() {
  const expireDays = Number(expireInput.value);
  if (withoutPassword()) return { noPassword: true, expireDays };
  const state = await readDraftState();
  return { password: passwordInput.value, expireDays, maxAccess: maxAccessValue(), jobId: state.jobId };
}

/** From the first link on, the choice (this password or none) applies to the whole mail. */
async function rememberShareChoice(jobId: string | null) {
  if (jobId === null) {
    await updateDraftState({ noPassword: true });
    lockNoPassword();
    return;
  }
  await updateDraftState({ jobId, password: passwordInput.value, maxAccess: maxAccessValue() });
  await tagDraft(jobId);
  lockPassword();
}

/**
 * Why a new link cannot be created with the current settings, or null. Upload links always need a password;
 * an own password must meet the rules.
 */
function settingsProblem(kind: "file" | "request"): string | null {
  if (kind === "request" && withoutPassword()) {
    return modeNone.disabled
      ? "Links zum Hochladen brauchen immer ein Kennwort. Diese Nachricht enthält bereits Links ohne Kennwort – " +
          "bitte dafür eine neue Nachricht verwenden."
      : "Links zum Hochladen brauchen immer ein Kennwort – bitte „Zufälliges Kennwort“ oder „Eigenes Kennwort“ wählen.";
  }
  if (!passwordInput.readOnly && !checkPassword(passwordInput.value).ok) {
    renderRules();
    return "Das Kennwort erfüllt noch nicht alle Anforderungen.";
  }
  return null;
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

/** Token of a share link (last path segment) – survives escaping or rewriting of the host in the body. */
const shareToken = (url: string) => url.replace(/\/+$/, "").split("/").pop() || url;

/**
 * Whether the body contains `marker` (a share URL). Retried briefly because the body may update asynchronously;
 * if the body cannot be read at all, the insertion is assumed to have worked.
 */
async function bodyContains(marker: string): Promise<boolean> {
  for (let attempt = 0; attempt < 3; attempt++) {
    let content: string;
    try {
      content = await officeAsync<string>((cb) => composeItem().body.getAsync(Office.CoercionType.Html, cb));
    } catch {
      return true;
    }
    if (content.includes(marker)) return true;
    await sleep(300);
  }
  return false;
}

/**
 * Inserts a card (HTML mail) or plain text (text mail) at the cursor. New Outlook reports success but inserts
 * nothing when the cursor has never been in the body (e.g. only recipients and subject were filled in) – so the
 * result is checked via `marker` and the card is put at the top of the mail instead. Throws if both fail.
 */
async function insertIntoBody(html: string, text: string, marker: string) {
  const body = composeItem().body;
  const type = await officeAsync<Office.CoercionType>((cb) => body.getTypeAsync(cb)).catch(
    () => Office.CoercionType.Html,
  );
  const isText = type === Office.CoercionType.Text;
  const data = isText ? text : html;
  const options = { coercionType: isText ? Office.CoercionType.Text : Office.CoercionType.Html };

  await officeAsync<void>((cb) => body.setSelectedDataAsync(data, options, cb));
  if (await bodyContains(marker)) return;

  await officeAsync<void>((cb) => body.prependAsync(data, options, cb));
  if (await bodyContains(marker)) return;

  throw new Error(
    "Outlook hat den Link nicht in die Nachricht übernommen. Bitte in den Nachrichtentext klicken und erneut " +
      "einfügen – den Link können Sie auch unter „Meine Freigaben“ kopieren.",
  );
}

/** Inserts the shared files as a card with download buttons (see mailCard.ts). */
async function insertLinks(results: UploadResult[]) {
  const files = results.map((r) => ({ name: r.fileName, size: r.size, url: r.shareUrl }));
  const expire = formatDate(results[0].expireDate);
  const withPassword = results[0].jobId !== null;
  await footerReady;
  await insertIntoBody(
    filesCard(files, expire, mailFooter, formatSize, withPassword),
    filesText(files, expire, mailFooter, formatSize, withPassword),
    shareToken(results[0].shareUrl),
  );
}

// --- Password: random (recommended), own or none -----------------------------------

const PASSWORD_HINT = "Die Empfänger erhalten das Kennwort nach dem Senden automatisch in einer separaten Nachricht.";
const NO_PASSWORD_HINT =
  "Jeder, der den Link kennt, kann die Dateien herunterladen. Es wird keine Nachricht mit Kennwort verschickt.";

/** "Ohne Kennwort" selected (file links only – upload links always keep a password). */
const withoutPassword = () => modeNone.checked;

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

/**
 * Random: generated, visible, read-only. Own: hidden input with "show" toggle and live rule check.
 * None: no password field, no "Kennwort abrufbar" – a plain link without password mail.
 */
function setPasswordMode(mode: "random" | "own" | "none") {
  const own = mode === "own";
  const none = mode === "none";
  modeRandom.checked = mode === "random";
  modeOwn.checked = own;
  modeNone.checked = none;
  passwordGroup.hidden = none;
  maxAccessField.hidden = none;
  passwordHint.textContent = none ? NO_PASSWORD_HINT : PASSWORD_HINT;
  passwordInput.readOnly = !own;
  generateButton.hidden = own;
  revealButton.hidden = !own;
  rulesList.hidden = !own;
  if (none) {
    passwordInput.value = "";
  } else if (own) {
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
modeNone.addEventListener("change", () => setPasswordMode("none"));
passwordInput.addEventListener("input", renderRules);
revealButton.addEventListener("click", () => setRevealed(passwordInput.type === "password"));

/** From the first upload on, the password applies to the whole mail and can no longer be changed. */
function lockPassword() {
  passwordInput.readOnly = true;
  modeRandom.disabled = true;
  modeOwn.disabled = true;
  modeNone.disabled = true;
  generateButton.hidden = true;
  rulesList.hidden = true;
  revealButton.hidden = false;
  setRevealed(false);
  passwordHint.textContent =
    "Dieses Kennwort gilt für alle Dateien dieser Nachricht. Die Empfänger erhalten es nach dem Senden automatisch.";
}

/** After the first link without password the whole mail stays without password (no mix in one mail). */
function lockNoPassword() {
  setPasswordMode("none");
  modeRandom.disabled = true;
  modeOwn.disabled = true;
  modeNone.disabled = true;
  passwordHint.textContent =
    "Die Links dieser Nachricht sind ohne Kennwort – jeder, der sie kennt, kann die Dateien herunterladen.";
}

/** Number of accesses within the allowed range (the server limits it again). */
function maxAccessValue(): number {
  const value = Math.floor(Number(maxAccessInput.value));
  if (!Number.isFinite(value) || value < 1) return serverConfig.sendMaxAccessCount;
  return Math.min(value, serverConfig.sendMaxAccessLimit);
}

/**
 * "Kennwort abrufbar" changed: before the first link it only applies to the next one; afterwards the mail's
 * Vaultwarden Send already exists and is updated on the server (until the password mail went out).
 */
maxAccessInput.addEventListener("change", async () => {
  const maxAccess = maxAccessValue();
  maxAccessInput.value = String(maxAccess);
  const state = await readDraftState();
  if (!state.jobId || state.maxAccess === maxAccess) return;
  try {
    const result = await api<{ maxAccess: number }>(
      `/api/jobs/${encodeURIComponent(state.jobId)}`,
      { method: "PATCH", body: JSON.stringify({ maxAccess }) },
      await getToken(),
    );
    maxAccessInput.value = String(result.maxAccess);
    await updateDraftState({ maxAccess: result.maxAccess });
    setStatus(
      `Der Link zum Kennwort ist jetzt ${result.maxAccess === 1 ? "einmal" : `${result.maxAccess}-mal`} abrufbar.`,
      "success",
    );
  } catch (err) {
    // Not changeable (e.g. links inserted before this version): show the value that actually applies.
    maxAccessInput.value = String(state.maxAccess ?? serverConfig.sendMaxAccessCount);
    if (err instanceof ApiError && err.code === "SEND_FIXED") maxAccessInput.readOnly = true;
    setStatus(`„Kennwort abrufbar“ konnte nicht geändert werden: ${errorMessage(err)}`, "error");
  }
});

submitButton.addEventListener("click", async () => {
  const selectedAttachments = Array.from(attachmentList.querySelectorAll<HTMLInputElement>("input:checked"));
  const selectedFiles = Array.from(pickedList.querySelectorAll<HTMLInputElement>("input:checked")).map(
    (box) => pickedFiles[Number(box.value)],
  );
  if (selectedAttachments.length + selectedFiles.length === 0) {
    return setStatus("Wählen Sie mindestens eine Datei aus.", "error");
  }
  const problem = settingsProblem("file");
  if (problem) {
    passwordInput.focus();
    return setStatus(problem, "error");
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
              (results[0].jobId === null
                ? `${results.length === 1 ? "Die Datei ist" : "Die Dateien sind"} ohne Kennwort abrufbar.`
                : `Das Kennwort wird nach dem Senden automatisch verschickt.`),
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

// --- Upload link ("Dateien von Empfängern erhalten") ----------------------------------------------

requestButton.addEventListener("click", async () => {
  // Upload links always need a password – otherwise anyone with the link could put files into the Nextcloud.
  const problem = settingsProblem("request");
  if (problem) {
    if (!withoutPassword()) passwordInput.focus();
    return setStatus(problem, "error");
  }
  requestButton.disabled = true;
  try {
    setStatus("Link zum Hochladen wird erstellt …");
    const result = await api<{ shareUrl: string; expireDate: string; jobId: string }>(
      "/api/requests",
      { method: "POST", body: JSON.stringify(await shareSettingsBody()) },
      await getToken(),
    );
    // Same password and password mail as the files of this mail.
    await rememberShareChoice(result.jobId);
    const expire = formatDate(result.expireDate);
    await footerReady;
    await insertIntoBody(
      requestCard(result.shareUrl, expire, mailFooter),
      requestText(result.shareUrl, expire, mailFooter),
      shareToken(result.shareUrl),
    );
    setStatus("Der Link zum Hochladen wurde eingefügt. Das Kennwort wird nach dem Senden automatisch verschickt.", "success");
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
  /** Password-protected link (Nextcloud cannot tell random from own passwords). */
  protected?: boolean;
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
    `<button type="button" class="btn btn-subtle insert">Als Link einfügen</button>` +
    `<button type="button" class="btn btn-subtle btn-danger revoke">Zurückziehen</button>` +
    `</div></div>`;
  row.querySelector<HTMLElement>(".share-name")!.textContent = share.name;
  row.querySelector<HTMLElement>(".share-name")!.title = share.name;
  row.querySelector<HTMLElement>(".share-meta")!.textContent =
    `Freigegeben am ${new Date(share.sharedAt).toLocaleDateString("de-DE")}` +
    (share.expiration ? ` · gültig bis ${formatDate(share.expiration)}` : "") +
    (share.protected === undefined ? "" : share.protected ? " · mit Kennwort" : " · ohne Kennwort");
  row.querySelector<HTMLElement>(".state-text")!.textContent = state.text;

  const copy = row.querySelector<HTMLButtonElement>(".copy")!;
  copy.addEventListener("click", async () => {
    await copyText(share.url);
    copy.textContent = "Kopiert";
    setTimeout(() => (copy.textContent = "Link kopieren"), 2000);
  });

  // A new link for this mail with the settings of the "Senden" tab – the original password cannot be recovered.
  const insert = row.querySelector<HTMLButtonElement>(".insert")!;
  insert.addEventListener("click", () => {
    linkAgain(share, insert).catch(() => undefined);
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

interface LinkAgainResult {
  kind: "file" | "request";
  shareUrl: string;
  expireDate: string;
  jobId: string | null;
  fileName: string;
  size: number | null;
}

/**
 * "Als Link einfügen" in "Meine Freigaben": creates a new link to the same file or upload folder with this mail's
 * settings (password and password mail, or "Ohne Kennwort" for files) and inserts it like a fresh one.
 */
async function linkAgain(share: ShareItem, button: HTMLButtonElement) {
  const problem = settingsProblem(share.kind);
  if (problem) {
    setStatus(`${problem} (Einstellungen unter „Senden“)`, "error");
    return;
  }
  button.disabled = true;
  try {
    setStatus(`Neuer Link zu „${share.name}“ wird erstellt …`);
    const result = await api<LinkAgainResult>(
      `/api/shares/${encodeURIComponent(share.id)}/link`,
      { method: "POST", body: JSON.stringify(await shareSettingsBody()) },
      await getToken(),
    );
    await rememberShareChoice(result.jobId);
    const expire = formatDate(result.expireDate);
    const withPassword = result.jobId !== null;
    await footerReady;
    if (result.kind === "request") {
      await insertIntoBody(
        requestCard(result.shareUrl, expire, mailFooter),
        requestText(result.shareUrl, expire, mailFooter),
        shareToken(result.shareUrl),
      );
    } else {
      const files = [{ name: result.fileName, size: result.size, url: result.shareUrl }];
      await insertIntoBody(
        filesCard(files, expire, mailFooter, formatSize, withPassword),
        filesText(files, expire, mailFooter, formatSize, withPassword),
        shareToken(result.shareUrl),
      );
    }
    setStatus(
      `Ein neuer Link zu „${share.name}“ wurde eingefügt. ` +
        (withPassword ? "Das Kennwort wird nach dem Senden automatisch verschickt." : "Er ist ohne Kennwort abrufbar."),
      "success",
    );
    loadShares().catch(() => undefined); // the new link shows up in the list
  } catch (err) {
    if (err instanceof ApiError && err.code === NOT_CONNECTED) showConnection({ connected: false });
    setStatus(`Der Link konnte nicht eingefügt werden: ${errorMessage(err)}`, "error");
  } finally {
    button.disabled = false;
  }
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
    setStatus("Große Anhänge werden bei dieser Nachricht normal mitgesendet.");
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
  } else if (state.noPassword) {
    lockNoPassword();
  } else {
    setPasswordMode("random");
  }
  optOutBox.checked = Boolean(state.optOut);
  // Under "enforce" there is no way to send large attachments without CloudShare.
  const enforce = serverConfig.attachmentPolicy === "enforce";
  $<HTMLLabelElement>("opt-out-row").hidden = enforce;
  $<HTMLParagraphElement>("opt-out-hint").hidden = enforce;
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
