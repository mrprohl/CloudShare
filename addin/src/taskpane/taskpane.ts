import {
  api,
  API_BASE,
  ApiError,
  composeItem,
  CONFIG_CACHE_KEY,
  getConfig,
  getFileAttachments,
  getToken,
  MB,
  NOT_CONNECTED,
  officeAsync,
  readDraftState,
  releaseDraftIfLinksGone,
  rememberLink,
  tagDraft,
  updateDraftState,
  type ServerConfig,
} from "../shared/office";
import { checkPassword, generatePassword } from "../shared/password";
import { locale, t } from "../shared/i18n";
import { errorText } from "../shared/sso";
import { EMPTY_PROFILE, fillFooter, hasPlaceholders, type Profile } from "./footer";
import { filesCard, filesText, messageCard, messageText, requestCard, requestText } from "./mailCard";

// All visible texts come from locales/<language>.json (t(), see shared/i18n.ts) and follow Outlook's terminology
// ("Nachricht"/"message", "Kennwort"/"password").

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
const modeNoneRow = $<HTMLLabelElement>("mode-none-row");
const passwordGroup = $<HTMLDivElement>("password-group");
const maxAccessField = $<HTMLDivElement>("max-access-field");
const passwordDaysField = $<HTMLDivElement>("password-days-field");
const passwordDaysInput = $<HTMLInputElement>("password-days");
const messageDaysInput = $<HTMLInputElement>("message-days");
const passwordToField = $<HTMLDivElement>("password-to-field");
const channelSelect = $<HTMLSelectElement>("password-channel");
const channelSmsOption = $<HTMLOptionElement>("channel-sms-option");
const passwordSmsInput = $<HTMLInputElement>("password-sms");
const passwordToInput = $<HTMLInputElement>("password-to");
const passwordToHint = $<HTMLParagraphElement>("password-to-hint");
const rulesList = $<HTMLUListElement>("password-rules");
const kindRadios = Array.from(document.querySelectorAll<HTMLInputElement>('input[name="kind"]'));
const filesSection = $<HTMLElement>("files-section");
const messageSection = $<HTMLElement>("message-section");
const requestSection = $<HTMLElement>("request-section");
const messageInput = $<HTMLTextAreaElement>("message");
const messageCount = $<HTMLParagraphElement>("message-count");
const validityToggle = $<HTMLButtonElement>("validity-toggle");
const validitySection = $<HTMLDivElement>("validity-section");
const validitySummary = $<HTMLSpanElement>("validity-summary");
const expireField = $<HTMLDivElement>("expire-field");
const messageDaysField = $<HTMLDivElement>("message-days-field");
const tabSend = $<HTMLButtonElement>("tab-send");
const tabShares = $<HTMLButtonElement>("tab-shares");
const panelSend = $<HTMLDivElement>("panel-send");
const panelShares = $<HTMLDivElement>("panel-shares");
const tabList = $<HTMLDivElement>("tabs");
const mailSection = $<HTMLElement>("mail-section");
const mailShares = $<HTMLDivElement>("mail-shares");
const mailNone = $<HTMLParagraphElement>("mail-none");
const sharesList = $<HTMLDivElement>("shares");
const noShares = $<HTMLParagraphElement>("no-shares");
const refreshSharesButton = $<HTMLButtonElement>("refresh-shares");
const expireInput = $<HTMLInputElement>("expire");
const maxAccessInput = $<HTMLInputElement>("max-access");
const submitButton = $<HTMLButtonElement>("submit");
const progress = $<HTMLDivElement>("progress");
const progressBar = $<HTMLDivElement>("progress-bar");
const statusBar = $<HTMLDivElement>("status");

/**
 * Opened from a mail being read (manifest: taskpane.html?view=read, e.g. in "Sent Items"): only "Meine Freigaben",
 * with the links found in that mail on top. Nothing is inserted or changed in the mail.
 */
const readMode = new URLSearchParams(location.search).get("view") === "read";

let serverConfig: ServerConfig;
/** MAIL_FOOTER with the placeholders filled from the sender's profile (see footer.ts). */
let mailFooter = "";
let footerReady: Promise<void> = Promise.resolve();
/** Files picked from the computer, not yet uploaded. */
let pickedFiles: File[] = [];
/** Uploads or insertions in progress – the check for deleted links waits meanwhile. */
let busy = 0;
/** A click on the footer button is being processed – it stays disabled meanwhile. */
let inserting = false;
/** Last known Nextcloud setting, re-applied when the password options are unlocked. */
let passwordEnforced = false;

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
    ? `${(bytes / 1024 ** 3).toLocaleString(locale(), { maximumFractionDigits: 1 })} GB`
    : bytes >= MB
      ? `${(bytes / MB).toLocaleString(locale(), { maximumFractionDigits: 1 })} MB`
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
  /** Nextcloud enforces passwords for public links – "Ohne Kennwort" is not possible then. */
  passwordEnforced?: boolean;
}

/** Storage behind the account: "[12,3 GB von 50 GB frei]", "[12,3 GB frei]" (no quota) or "[3,2 GB belegt]". */
function storageText(status: ConnectionStatus): string {
  const isBytes = (value: unknown): value is number => typeof value === "number" && value >= 0;
  if (isBytes(status.freeBytes) && isBytes(status.totalBytes)) {
    return ` [${t("storage.freeOf", { free: formatSize(status.freeBytes), total: formatSize(status.totalBytes) })}]`;
  }
  if (isBytes(status.freeBytes)) return ` [${t("storage.free", { free: formatSize(status.freeBytes) })}]`;
  if (isBytes(status.usedBytes)) return ` [${t("storage.used", { used: formatSize(status.usedBytes) })}]`;
  return "";
}

function showConnection(status: ConnectionStatus) {
  connectView.hidden = status.connected;
  sendView.hidden = !status.connected;
  connectHint.hidden = true;
  accountText.textContent = `${status.loginName ?? ""}${storageText(status)}`;
  applyPasswordPolicy(status.passwordEnforced === true);
  // Reading a mail: the list is all there is – load it as soon as the connection is there (also right after connecting).
  if (readMode && status.connected) loadShares().catch((err) => setStatus(errorMessage(err), "error"));
}

/** Hides "Ohne Kennwort" when Nextcloud enforces link passwords (it would refuse such a link). */
function applyPasswordPolicy(enforced: boolean) {
  passwordEnforced = enforced;
  // Switch back unless this mail already contains links without password.
  if (enforced && modeNone.checked && !modeNone.disabled) setPasswordMode("random");
  renderKind();
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
        setStatus(t("connect.done"), "success");
        return;
      }
    }
    if (attempt === connectAttempt) {
      connectHint.hidden = true;
      setStatus(t("connect.incomplete"), "error");
    }
  } catch (err) {
    popup?.close();
    connectHint.hidden = true;
    setStatus(t("connect.failed", { error: errorMessage(err) }), "error");
  } finally {
    connectButton.disabled = false;
  }
});

$<HTMLAnchorElement>("disconnect").addEventListener("click", async (e) => {
  e.preventDefault();
  try {
    await api("/api/nextcloud/connection", { method: "DELETE" }, await getToken());
    showConnection({ connected: false });
    setStatus(t("disconnect.done"));
  } catch (err) {
    setStatus(t("disconnect.failed", { error: errorMessage(err) }), "error");
  }
});

// --- File lists ------------------------------------------------------------------

const FILE_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M5.5 2.5h5.3l3.7 3.7v10.3a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-13a1 1 0 0 1 1-1Z" fill="none" stroke="currentColor"/>' +
  '<path d="M10.5 2.5v4h4" fill="none" stroke="currentColor"/></svg>';

/** One row in Outlook's list style: checkbox, file icon, name, size and an optional "Groß" badge. */
function fileRow(
  name: string,
  size: number,
  value: string,
  checked: boolean,
  large: boolean,
  /** Files picked from the computer can be removed from the list again (mail attachments are removed in Outlook). */
  onRemove?: () => void,
): HTMLLabelElement {
  const row = document.createElement("label");
  row.className = "checkbox list-item";
  row.innerHTML =
    `<input type="checkbox" /><span class="box" aria-hidden="true"></span>${FILE_ICON}` +
    `<span class="file-text"><span class="file-name"></span><span class="file-size"></span></span>` +
    (large ? `<span class="badge"></span>` : "");
  const box = row.querySelector("input")!;
  box.value = value;
  box.checked = checked;
  box.dataset.name = name;
  row.querySelector(".file-name")!.textContent = name;
  row.querySelector<HTMLElement>(".file-name")!.title = name;
  row.querySelector(".file-size")!.textContent = formatSize(size);
  const badge = row.querySelector<HTMLElement>(".badge");
  if (badge) badge.textContent = t("files.large");
  // Own name for the checkbox: otherwise a remove button inside the label would become part of it.
  box.setAttribute("aria-label", `${name}, ${formatSize(size)}`);
  if (onRemove) {
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "btn btn-icon btn-subtle remove-file";
    remove.title = t("files.remove", { name });
    remove.setAttribute("aria-label", t("files.removeLabel", { name }));
    remove.innerHTML =
      `<svg class="icon" viewBox="0 0 20 20" aria-hidden="true"><path d="M5.5 5.5l9 9m0-9-9 9" fill="none" ` +
      `stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
    remove.addEventListener("click", (event) => {
      // Inside the row's label: do not toggle the checkbox.
      event.preventDefault();
      event.stopPropagation();
      onRemove();
    });
    row.append(remove);
  }
  return row;
}

/** Ticked files (attachments and files from the computer) – the button says how many get a link. */
const selectedFileCount = () =>
  attachmentList.querySelectorAll("input:checked").length + pickedList.querySelectorAll("input:checked").length;

async function renderAttachments() {
  const attachments = await getFileAttachments();
  const threshold = serverConfig.thresholdMb * MB;
  // Large attachments are preselected.
  attachmentList.replaceChildren(
    ...attachments.map((a) => fileRow(a.name, a.size, a.id, a.size > threshold, a.size > threshold)),
  );
  noAttachments.hidden = attachments.length > 0 || pickedFiles.length > 0;
  renderSubmit();
}

function renderPicked() {
  // Re-rendering renumbers the rows – keep the files the user unticked unticked.
  const unticked = new Set(
    Array.from(pickedList.querySelectorAll<HTMLInputElement>("input:not(:checked)")).map((box) => pickedRendered[Number(box.value)]),
  );
  pickedRendered = pickedFiles;
  pickedList.replaceChildren(
    ...pickedFiles.map((file, i) =>
      fileRow(file.name, file.size, String(i), !unticked.has(file), false, () => {
        pickedFiles = pickedFiles.filter((f) => f !== file);
        renderPicked();
      }),
    ),
  );
  noAttachments.hidden = attachmentList.childElementCount > 0 || pickedFiles.length > 0;
  renderSubmit();
}

/** The files the rows of #picked currently show (their checkbox values are indexes into it). */
let pickedRendered: File[] = [];

for (const list of [attachmentList, pickedList]) list.addEventListener("change", () => renderSubmit());


/**
 * Diagnostics of the taskpane (only with CLIENT_TRACE=true): reports a step to the server log as a plain GET, like
 * the send handler does. Used for the file dialog, which classic Outlook sometimes does not open.
 */
const paneRun = Math.random().toString(36).slice(2, 10);
function paneTrace(step: string, detail = "") {
  if (!serverConfig?.trace) return;
  const host = Office.context.diagnostics ? `${Office.context.diagnostics.platform} ${Office.context.diagnostics.version}` : "?";
  const query = [["run", paneRun], ["step", step], ["host", host], ["origin", location.origin], ["detail", detail]]
    .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
    .join("&");
  fetch(`${API_BASE}/api/trace?${query}`).catch(() => undefined);
}

type OpenFilePicker = (options: { multiple: boolean }) => Promise<Array<{ getFile(): Promise<File> }>>;

/**
 * "Dateien vom Computer hinzufügen": first the native dialog of the File System Access API (Chromium, Edge, WebView2 –
 * a different path than a file input), then the hidden file input as fallback. Classic Outlook's WebView did not
 * reliably open the dialog for the file input (script click, label and transparent overlay were all tried).
 */
async function openFilePicker() {
  const showOpenFilePicker = (window as unknown as { showOpenFilePicker?: OpenFilePicker }).showOpenFilePicker;
  paneTrace("picker:click", showOpenFilePicker ? "fsa" : "input");
  if (showOpenFilePicker) {
    try {
      const handles = await showOpenFilePicker.call(window, { multiple: true });
      const files = await Promise.all(handles.map((handle) => handle.getFile()));
      paneTrace("picker:chosen", String(files.length));
      addPickedFiles(files);
      return;
    } catch (err) {
      const name = err instanceof Error ? err.name : String(err);
      paneTrace("picker:fsa-error", `${name}: ${err instanceof Error ? err.message : ""}`.slice(0, 200));
      if (name === "AbortError") return; // the user closed the dialog
      // SecurityError / NotAllowedError etc.: try the file input instead.
    }
  }
  paneTrace("picker:input");
  fileInput.click();
}

// The whole drop zone opens the dialog; a click on its button ("vom Computer auswählen") bubbles up to it, so the
// dialog opens exactly once (the button keeps it reachable by keyboard).
$<HTMLDivElement>("dropzone").addEventListener("click", () => void openFilePicker());

fileInput.addEventListener("change", () => {
  paneTrace("picker:change", String(fileInput.files?.length ?? 0));
  addPickedFiles(Array.from(fileInput.files ?? []));
  fileInput.value = "";
});
fileInput.addEventListener("cancel", () => paneTrace("picker:cancel"));

function addPickedFiles(files: File[]) {
  if (files.length === 0) return;
  pickedFiles = [...pickedFiles, ...files];
  renderPicked();
}

// Drag & drop from Explorer / desktop onto the "Senden" tab – needs no file dialog. In classic Outlook (WebView2)
// the dialog sometimes opens behind the Outlook window, which looks like "nothing happens".
const dropZone = $<HTMLDivElement>("panel-send");
const hasFiles = (event: DragEvent) => Array.from(event.dataTransfer?.types ?? []).includes("Files");
let dragDepth = 0;
dropZone.addEventListener("dragenter", (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  // Files dragged in while "Nachricht" or "Anfordern" is shown: show the file list and its drop box.
  if (selectedKind() !== "files") selectKind("files");
  dragDepth++;
  dropZone.classList.add("dropping");
});
dropZone.addEventListener("dragover", (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault(); // allows the drop
  event.dataTransfer!.dropEffect = "copy";
});
dropZone.addEventListener("dragleave", () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) dropZone.classList.remove("dropping");
});
dropZone.addEventListener("drop", (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault(); // otherwise the WebView would open the file itself
  if (selectedKind() !== "files") selectKind("files");
  dragDepth = 0;
  dropZone.classList.remove("dropping");
  // Folders cannot be uploaded – they arrive as entries without type and size 0 (and fail to read).
  const files = Array.from(event.dataTransfer?.files ?? []).filter((f) => f.size > 0 || f.type !== "");
  addPickedFiles(files);
  if (files.length > 0) {
    setStatus(t("files.added", { count: files.length }), "success");
  }
});
// Dropping a file next to the drop zone must not navigate the taskpane away.
window.addEventListener("dragover", (event) => event.preventDefault());
window.addEventListener("drop", (event) => event.preventDefault());

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
    throw new Error(t("files.unsupported"));
  }
  return base64ToBlob(content.content);
}

// --- Upload and share ---------------------------------------------------------

interface UploadStart { uploadId: string; chunkSize: number; totalChunks: number }
/** jobId is null for links "Ohne Kennwort" – no password mail. */
interface UploadResult {
  shareUrl: string;
  expireDate: string;
  /** Password job; for links "Ohne Kennwort" a job for the download confirmation, or null. */
  jobId: string | null;
  /** Link without password (older servers: recognisable by jobId null only). */
  noPassword?: boolean;
  fileName: string;
  size: number;
}

/** Link "Ohne Kennwort"? */
const isWithoutPassword = (result: { jobId: string | null; noPassword?: boolean }) =>
  result.noPassword === true || result.jobId === null;

async function upload(blob: Blob, fileName: string, step: string): Promise<UploadResult> {
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
      setStatus(t("files.uploading", { name: fileName, percent }));
    }
  } catch (err) {
    await api(`/api/uploads/${start.uploadId}`, { method: "DELETE" }, await getToken()).catch(() => undefined);
    throw err;
  }

  const result = await api<UploadResult>(
    `/api/uploads/${start.uploadId}/complete`,
    { method: "POST", body: JSON.stringify(await shareSettingsBody(step)) },
    await getToken(),
  );
  await rememberShareChoice(result.jobId, isWithoutPassword(result));
  return result;
}

/**
 * Settings of the "Senden" tab for a new link. `step` identifies the click ("Als Link einfügen" etc.): every step
 * gets its own password and password mail, files of the same step share them (see attachToJob on the server).
 */
async function shareSettingsBody(step: string) {
  const expireDays = Number(expireInput.value);
  if (withoutPassword()) return { noPassword: true, expireDays, jobId: (await readDraftState()).jobId };
  // An address or number typed just before the click (its "change" event may not have run yet) must reach the job.
  const wanted = channelTarget();
  const current = await readDraftState();
  if (wanted && (current.passwordTo !== wanted.passwordTo || current.passwordSms !== wanted.passwordSms)) {
    await updateDraftState(wanted);
  }
  const state = await readDraftState();
  return {
    password: passwordInput.value,
    expireDays,
    maxAccess: maxAccessValue(),
    // By SMS the password is in the SMS itself; the password link only holds it until sending – as long as possible.
    passwordDays: selectedChannel() === "sms" ? sendMaxDays() : sendDays(passwordDaysInput).value,
    jobId: state.jobId,
    step,
  };
}

/** New ID for one click on "Als Link einfügen" / "Link zum Hochladen einfügen" / "Meine Freigaben → Als Link einfügen". */
const newStep = () => Math.random().toString(36).slice(2, 12);

/** From the first link on, the choice (this password or none) applies to the whole mail. */
async function rememberShareChoice(jobId: string | null, noPassword = jobId === null) {
  if (noPassword || jobId === null) {
    // Links without password. With the download confirmation the server keeps a job without password, which the
    // send handler arms like a password job (it only waits for the sent mail).
    await updateDraftState(jobId ? { noPassword: true, jobId } : { noPassword: true });
    if (jobId) await tagDraft(jobId);
    lockNoPassword();
    return;
  }
  let state = await readDraftState();
  if (state.password) {
    // Mail started before passwords per step existed: it keeps its one shared password.
    state = await updateDraftState({ jobId });
    await tagDraft(jobId);
  } else {
    // One password per step: only the mode is fixed; the next step gets a new password (nextStepPassword).
    state = await updateDraftState({ jobId, separate: true });
    await tagDraft(jobId);
    lockSeparate();
  }
  // Chosen before the first link: hand the channel to the new job (the send handler passes it on again when arming).
  if (state.passwordTo || state.passwordSms) {
    await api(`/api/jobs/${encodeURIComponent(jobId)}`, {
      method: "PATCH",
      body: JSON.stringify(state.passwordSms ? { passwordSms: state.passwordSms } : { passwordTo: state.passwordTo }),
    }, await getToken()).catch(() => undefined);
  }
}

// --- "Kennwort senden an": recipients of the mail / another mail address / SMS ------------------------------

const looksLikeEmail = (value: string) => /^[^\s@,;<>"()]+@[^\s@,;<>"()]+\.[^\s@,;<>"()]{2,}$/.test(value);

/** Same rules as the server (server/src/sms.ts): E.164, "00" → "+", leading "0" → default country code. */
function normalizePhone(input: string): string | null {
  let number = input.trim().replace(/[\s\-./()]/g, "");
  if (number.startsWith("00")) number = `+${number.slice(2)}`;
  else if (/^0[1-9]/.test(number)) number = `+${serverConfig.smsDefaultCountry ?? "49"}${number.slice(1)}`;
  return /^\+[1-9]\d{7,14}$/.test(number) ? number : null;
}

type Channel = "recipients" | "mail" | "sms";
const selectedChannel = (): Channel => {
  const value = channelSelect.value;
  return value === "sms" || value === "mail" ? value : "recipients";
};

const channelHint = (channel: Channel) =>
  channel === "mail" ? t("channel.hintMail") : channel === "sms" ? t("channel.hintSms") : "";

function showChannel(channel: Channel) {
  channelSelect.value = channel;
  passwordToInput.hidden = channel !== "mail";
  passwordSmsInput.hidden = channel !== "sms";
  passwordToHint.hidden = channel === "recipients";
  passwordToHint.textContent = channelHint(channel);
  renderValidity();
}

/** The valid target of the chosen channel ({} = recipients), or null while the address / number is not valid. */
function channelTarget(): { passwordTo?: string; passwordSms?: string } | null {
  const channel = selectedChannel();
  if (channel === "mail") {
    const address = passwordToInput.value.trim();
    return looksLikeEmail(address) ? { passwordTo: address, passwordSms: undefined } : null;
  }
  if (channel === "sms") {
    const number = normalizePhone(passwordSmsInput.value);
    return number ? { passwordTo: undefined, passwordSms: number } : null;
  }
  return { passwordTo: undefined, passwordSms: undefined };
}

/** Stores the channel for this mail and, once the job exists, on the server. */
async function applyChannel(target: { passwordTo?: string; passwordSms?: string }) {
  const state = await updateDraftState(target);
  try {
    if (state.jobId) {
      await api(`/api/jobs/${encodeURIComponent(state.jobId)}`, {
        method: "PATCH",
        body: JSON.stringify(target.passwordSms ? { passwordSms: target.passwordSms } : { passwordTo: target.passwordTo ?? "" }),
      }, await getToken());
    }
    setStatus(
      target.passwordSms
        ? t("channel.savedSms", { to: target.passwordSms })
        : target.passwordTo
          ? t("channel.savedMail", { to: target.passwordTo })
          : t("channel.savedRecipients"),
      "success",
    );
  } catch (err) {
    setStatus(t("channel.saveFailed", { error: errorMessage(err) }), "error");
  }
}

channelSelect.addEventListener("change", () => {
  const channel = selectedChannel();
  showChannel(channel);
  if (channel === "mail") passwordToInput.focus();
  if (channel === "sms") passwordSmsInput.focus();
  const target = channelTarget();
  if (target) void applyChannel(target);
});

passwordToInput.addEventListener("change", () => {
  const value = passwordToInput.value.trim();
  passwordToInput.value = value;
  if (value && !looksLikeEmail(value)) {
    passwordToInput.setAttribute("aria-invalid", "true");
    setStatus(t("channel.mailInvalid"), "error");
    return;
  }
  passwordToInput.removeAttribute("aria-invalid");
  if (value) void applyChannel({ passwordTo: value, passwordSms: undefined });
});

passwordSmsInput.addEventListener("change", () => {
  const value = passwordSmsInput.value.trim();
  if (!value) return;
  const number = normalizePhone(value);
  if (!number) {
    passwordSmsInput.setAttribute("aria-invalid", "true");
    setStatus(t("channel.smsInvalid"), "error");
    return;
  }
  passwordSmsInput.removeAttribute("aria-invalid");
  passwordSmsInput.value = number;
  void applyChannel({ passwordTo: undefined, passwordSms: number });
});

/**
 * Why a new link cannot be created with the current settings, or null. Upload links always need a password;
 * an own password must meet the rules; the chosen password channel needs its address or number.
 */
function settingsProblem(kind: "file" | "request" | "message"): string | null {
  if (kind !== "file" && withoutPassword()) {
    return modeNone.disabled ? t(`${kind}.needsPasswordLocked`) : t(`${kind}.needsPassword`);
  }
  if (!passwordInput.readOnly && !checkPassword(passwordInput.value).ok) {
    renderRules();
    passwordInput.focus();
    return t("password.weak");
  }
  // Another address / SMS chosen but not filled in: the password would silently go to the recipients instead – so no
  // link until it is there (or "an die Empfänger der Nachricht" is chosen).
  if (!withoutPassword() && channelTarget() === null) {
    const sms = selectedChannel() === "sms";
    const input = sms ? passwordSmsInput : passwordToInput;
    input.setAttribute("aria-invalid", "true");
    input.focus();
    if (sms) return t(input.value.trim() ? "channel.smsInvalidShort" : "channel.smsMissing");
    return t(input.value.trim() ? "channel.mailInvalidShort" : "channel.mailMissing");
  }
  return null;
}

const formatDate = (isoDate: string) => new Date(isoDate).toLocaleDateString(locale());

/** Empty paragraph after an inserted card, so the cursor ends up below it (see insertIntoBody). */
const AFTER_CARD = "<p style=\"margin:0;\"><br></p>";

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
  // An empty line after the card: Outlook leaves the cursor at the end of the inserted HTML – without it the cursor
  // stays in the card's last cell, and the next "Als Link einfügen" would land inside this card (broken layout).
  const data = isText ? text : `${html}${AFTER_CARD}`;
  const options = { coercionType: isText ? Office.CoercionType.Text : Office.CoercionType.Html };

  // Only links that really landed in the mail are recorded: after a failed insertion the user may paste the link
  // by hand ("Meine Freigaben"), so the draft must not be released then (see releaseDraftIfLinksGone).
  await officeAsync<void>((cb) => body.setSelectedDataAsync(data, options, cb));
  if (await bodyContains(marker)) return rememberLink(marker);

  await officeAsync<void>((cb) => body.prependAsync(data, options, cb));
  if (await bodyContains(marker)) return rememberLink(marker);

  throw new Error(t("insert.notTaken"));
}

/** Inserts the shared files as a card with download buttons (see mailCard.ts). */
async function insertLinks(results: UploadResult[]) {
  const files = results.map((r) => ({ name: r.fileName, size: r.size, url: r.shareUrl }));
  const expire = formatDate(results[0].expireDate);
  const withPassword = !isWithoutPassword(results[0]);
  await footerReady;
  await insertIntoBody(
    filesCard(files, expire, mailFooter, formatSize, withPassword),
    filesText(files, expire, mailFooter, formatSize, withPassword),
    shareToken(results[0].shareUrl),
  );
}

// --- Password: random (recommended), own or none -----------------------------------


/** "Ohne Kennwort" selected (file links only – upload links always keep a password). */
const withoutPassword = () => modeNone.checked;

function setRevealed(revealed: boolean) {
  passwordInput.type = revealed ? "text" : "password";
  revealButton.setAttribute("aria-pressed", String(revealed));
  const label = t(revealed ? "password.hide" : "password.show");
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
  passwordToField.hidden = none;
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
  renderPasswordHint();
  renderValidity();
}

/** The mail already has links without password: a confidential message or upload link cannot join it. */
const blockedByNoPassword = () => selectedKind() !== "files" && modeNone.checked && modeNone.disabled;

/** The line below the password: what happens with it in this mail. */
function renderPasswordHint() {
  const kind = selectedKind();
  let key: string;
  if (blockedByNoPassword()) key = `${kind}.needsPasswordLocked`;
  else if (modeNone.checked) key = modeNone.disabled ? "password.hintLockedNone" : "password.hintNone";
  else if (modeRandom.disabled) key = "password.hintLocked"; // one password for the whole mail (older drafts)
  else if (perStep) key = "password.hintSeparate";
  else key = "password.hint";
  passwordHint.textContent = t(key);
  passwordHint.classList.toggle("warning", blockedByNoPassword());
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
  renderPasswordHint();
}

/** The mail has links with a password per step: every further click gets a new password. */
let perStep = false;

/**
 * After the first link with a password: every "Als Link einfügen" gets its own password and password mail, so the
 * password stays open for the next step – only "Ohne Kennwort" is no longer possible in this mail.
 */
function lockSeparate() {
  perStep = true;
  modeNone.disabled = true;
  renderPasswordHint();
}

/** After a step: a new random password for the next one, or an empty field for an own one. */
function nextStepPassword() {
  if (!perStep) return;
  if (modeRandom.checked) {
    passwordInput.value = generatePassword();
  } else if (modeOwn.checked) {
    passwordInput.value = "";
    renderRules();
  }
}

/** After the first link without password the whole mail stays without password (no mix in one mail). */
function lockNoPassword() {
  setPasswordMode("none");
  modeRandom.disabled = true;
  modeOwn.disabled = true;
  modeNone.disabled = true;
  renderKind();
}

/** Password options as for a new mail (after all CloudShare links were deleted from it). */
function unlockPassword() {
  modeRandom.disabled = false;
  modeOwn.disabled = false;
  modeNone.disabled = false;
  perStep = false;
  maxAccessInput.readOnly = false;
  maxAccessInput.value = String(serverConfig.sendMaxAccessCount);
  passwordDaysInput.value = String(sendDefaultDays());
  setPasswordMode("random");
  applyPasswordPolicy(passwordEnforced);
}

/**
 * The user may delete the inserted links from the mail again. Then the password choice no longer applies and must
 * not stay locked, and no password mail must go out (the send handler checks this too).
 */
async function releaseIfLinksDeleted() {
  // Only while the mail has CloudShare links (options locked or per-file mode fixed) and nothing is running.
  if (busy > 0 || document.hidden || !(modeRandom.disabled || perStep)) return;
  if (await releaseDraftIfLinksGone().catch(() => false)) {
    unlockPassword();
    setStatus(t("password.released"));
  }
}

/** Number of accesses within the allowed range (the server limits it again). */
function maxAccessValue(): number {
  return boundedNumber(maxAccessInput, serverConfig.sendMaxAccessCount, serverConfig.sendMaxAccessLimit).value;
}

/** Lifetime of a Vaultwarden Send (password link, confidential message): default and limit from the server. */
const sendDefaultDays = () => serverConfig?.sendExpireDays ?? 5;
const sendMaxDays = () => serverConfig?.sendMaxExpireDays ?? 30;
const sendDays = (input: HTMLInputElement) => boundedNumber(input, sendDefaultDays(), sendMaxDays());

/** Whole number from 1 to `max` (empty or invalid → `fallback`); `capped` if the entry was above the maximum. */
function boundedNumber(input: HTMLInputElement, fallback: number, max: number): { value: number; capped: boolean } {
  const raw = Math.floor(Number(input.value));
  if (!Number.isFinite(raw) || raw < 1) return { value: fallback, capped: false };
  return raw > max ? { value: max, capped: true } : { value: raw, capped: false };
}

const timesText = (count: number) => t("maxAccess.times", { count });
/** Sentences of a status message, joined with a space (empty ones left out). */
const sentences = (...parts: string[]) => parts.filter(Boolean).join(" ");

/** "Link gültig für": applies to the next links; above the maximum it is capped with a note (like the server does). */
expireInput.addEventListener("change", async () => {
  const { value, capped } = boundedNumber(expireInput, serverConfig.defaultExpireDays, serverConfig.maxExpireDays);
  expireInput.value = String(value);
  const state = await readDraftState();
  setStatus(
    sentences(
      capped ? t("expire.capped", { max: serverConfig.maxExpireDays }) : "",
      t("expire.changed", { count: value }),
      state.links?.length ? t("expire.keepExisting") : "",
    ),
    capped ? "info" : "success",
  );
});

/**
 * "Kennwort-Link gültig für" / "Nachricht gültig für": apply to the next password link or message (the Sends of links
 * already inserted keep their date); above the maximum the value is capped with a note.
 */
for (const [input, changed] of [
  [passwordDaysInput, "passwordDays.changed"],
  [messageDaysInput, "message.daysChanged"],
] as const) {
  input.addEventListener("change", () => {
    const { value, capped } = sendDays(input);
    input.value = String(value);
    setStatus(
      sentences(capped ? t("expire.capped", { max: sendMaxDays() }) : "", t(changed, { count: value })),
      capped ? "info" : "success",
    );
  });
}

/**
 * "Kennwort abrufbar" changed: before the first link it only applies to the next one; afterwards the mail's
 * Vaultwarden Send already exists and is updated on the server (until the password mail went out).
 */
maxAccessInput.addEventListener("change", async () => {
  const { value: maxAccess, capped } = boundedNumber(
    maxAccessInput,
    serverConfig.sendMaxAccessCount,
    serverConfig.sendMaxAccessLimit,
  );
  maxAccessInput.value = String(maxAccess);
  const note = capped ? t("maxAccess.capped", { max: serverConfig.sendMaxAccessLimit }) : "";
  const state = await readDraftState();
  // No link yet, or one password per step: the number applies to the next link, nothing to change on the server.
  if (!state.jobId || state.separate || state.maxAccess === maxAccess) {
    setStatus(
      sentences(note, t(state.separate ? "maxAccess.nextStep" : "maxAccess.next", { times: timesText(maxAccess) })),
      capped ? "info" : "success",
    );
    return;
  }
  try {
    const result = await api<{ maxAccess: number }>(
      `/api/jobs/${encodeURIComponent(state.jobId)}`,
      { method: "PATCH", body: JSON.stringify({ maxAccess }) },
      await getToken(),
    );
    maxAccessInput.value = String(result.maxAccess);
    renderValidity();
    await updateDraftState({ maxAccess: result.maxAccess });
    setStatus(sentences(note, t("maxAccess.now", { times: timesText(result.maxAccess) })), capped ? "info" : "success");
  } catch (err) {
    // Not changeable (e.g. links inserted before this version): show the value that actually applies.
    maxAccessInput.value = String(state.maxAccess ?? serverConfig.sendMaxAccessCount);
    renderValidity();
    if (err instanceof ApiError && err.code === "SEND_FIXED") maxAccessInput.readOnly = true;
    setStatus(t("maxAccess.failed", { error: errorMessage(err) }), "error");
  }
});

generateButton.addEventListener("click", () => {
  passwordInput.value = generatePassword();
});

// --- What to insert: "Dateien" / "Nachricht" / "Anfordern" ------------------------------------------------------
// One kind at a time and one button in the footer that says what it inserts. Password, password channel and validity
// below apply to all three (every click is one step with its own password, see shareSettingsBody).

type Kind = "files" | "message" | "request";
const selectedKind = (): Kind => {
  const value = kindRadios.find((radio) => radio.checked)?.value;
  return value === "message" || value === "request" ? value : "files";
};

function selectKind(kind: Kind) {
  for (const radio of kindRadios) radio.checked = radio.value === kind;
  renderKind();
}

for (const radio of kindRadios) radio.addEventListener("change", () => renderKind());

/**
 * Shows the part of the chosen kind. "Ohne Kennwort" exists only for files – a confidential message or an upload link
 * always needs a password, so the option is not offered there (instead of an error after the click).
 */
function renderKind() {
  const kind = selectedKind();
  filesSection.hidden = kind !== "files";
  messageSection.hidden = kind !== "message";
  requestSection.hidden = kind !== "request";
  // Chosen for files but not fixed yet: a text or upload link gets a random password instead.
  if (kind !== "files" && modeNone.checked && !modeNone.disabled) setPasswordMode("random");
  const lockedNone = modeNone.checked && modeNone.disabled;
  // A mail fixed to "Ohne Kennwort" keeps showing it, so the note below explains why nothing else can be inserted.
  modeNoneRow.hidden = !lockedNone && (passwordEnforced || kind !== "files");
  renderPasswordHint();
  renderValidity();
  renderSubmit();
}

/** The footer button names what it inserts ("2 Dateien als Link einfügen", "Vertrauliche Nachricht einfügen", …). */
function renderSubmit() {
  const kind = selectedKind();
  if (kind === "message") {
    submitButton.textContent = t("message.button");
  } else if (kind === "request") {
    submitButton.textContent = t("request.button");
  } else {
    const count = selectedFileCount();
    submitButton.textContent = count > 0 ? t("submit.files", { count }) : t("submit.button");
  }
  submitButton.disabled = inserting || blockedByNoPassword();
}

/** "Gültigkeit": only the fields of the chosen kind, collapsed behind a one-line summary of the values. */
function renderValidity() {
  if (!serverConfig) return;
  const kind = selectedKind();
  // "Kennwort abrufbar" / "Kennwort-Link gültig für" only for a password link: not without password, not by SMS (the
  // password is in the SMS itself and works as long as the link does).
  const passwordLink = !withoutPassword() && selectedChannel() !== "sms";
  expireField.hidden = kind === "message";
  messageDaysField.hidden = kind !== "message";
  maxAccessField.hidden = !passwordLink;
  passwordDaysField.hidden = !passwordLink;
  const parts = [
    kind === "message"
      ? t("validity.message", { count: sendDays(messageDaysInput).value })
      : t("validity.link", {
          count: boundedNumber(expireInput, serverConfig.defaultExpireDays, serverConfig.maxExpireDays).value,
        }),
  ];
  if (passwordLink) {
    parts.push(t("validity.password", { count: sendDays(passwordDaysInput).value, times: timesText(maxAccessValue()) }));
  }
  validitySummary.textContent = parts.join(" · ");
}

validityToggle.addEventListener("click", () => {
  const open = validitySection.hidden === true;
  validityToggle.setAttribute("aria-expanded", String(open));
  validitySection.hidden = !open;
});
// After the fields' own handlers (they cap the value first).
validitySection.addEventListener("change", () => renderValidity());

submitButton.addEventListener("click", () => {
  const kind = selectedKind();
  void (kind === "message" ? insertMessage() : kind === "request" ? insertRequest() : insertFiles());
});

/** One click on the button: disabled until done, then the next step gets its own password. */
async function runInsert(work: () => Promise<void>) {
  inserting = true;
  renderSubmit();
  busy++;
  try {
    await work();
  } finally {
    inserting = false;
    nextStepPassword();
    busy--;
    renderSubmit();
  }
}

// --- Files ---------------------------------------------------------------------------------

async function insertFiles() {
  const selectedAttachments = Array.from(attachmentList.querySelectorAll<HTMLInputElement>("input:checked"));
  const selectedFiles = Array.from(pickedList.querySelectorAll<HTMLInputElement>("input:checked")).map(
    (box) => pickedFiles[Number(box.value)],
  );
  if (selectedAttachments.length + selectedFiles.length === 0) {
    return setStatus(t("submit.noneSelected"), "error");
  }
  const problem = settingsProblem("file");
  if (problem) {
    return setStatus(problem, "error");
  }

  await runInsert(async () => {
    const results: UploadResult[] = [];
    // One step for this click: all selected files share one password and one password mail.
    const step = newStep();
    let failed = false;
    try {
      for (const box of selectedAttachments) {
        const name = box.dataset.name ?? t("files.attachment");
        setProgress(0);
        setStatus(t("files.preparing", { name }));
        const blob = await readAttachment(box.value);
        results.push(await upload(blob, name, step));
        // Remove from the mail only after a successful upload.
        await officeAsync<void>((cb) => composeItem().removeAttachmentAsync(box.value, cb));
      }
      for (const file of selectedFiles) {
        if (file.size === 0) throw new Error(t("files.empty", { name: file.name }));
        setProgress(0);
        results.push(await upload(file, file.name, step));
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
              sentences(
                t("submit.inserted", { count: results.length }),
                isWithoutPassword(results[0])
                  ? t("submit.noPassword", { count: results.length })
                  : t("submit.passwordFollows"),
              ),
              "success",
            );
          }
        } catch (err) {
          setStatus(t("insert.failed", { error: errorMessage(err) }), "error");
        }
        renderPicked();
        await renderAttachments().catch(() => undefined);
        await refreshConnection().catch(() => undefined); // refreshes the free quota
      }
      setProgress(null);
    }
  });
}

// --- Upload link ("Dateien von Empfängern erhalten") ----------------------------------------------

async function insertRequest() {
  // Upload links always need a password – otherwise anyone with the link could put files into the Nextcloud.
  const problem = settingsProblem("request");
  if (problem) {
    return setStatus(problem, "error");
  }
  await runInsert(async () => {
    try {
      setStatus(t("request.creating"));
      const result = await api<{ shareUrl: string; expireDate: string; jobId: string }>(
        "/api/requests",
        { method: "POST", body: JSON.stringify(await shareSettingsBody(newStep())) },
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
      setStatus(t("request.inserted"), "success");
    } catch (err) {
      if (err instanceof ApiError && err.code === NOT_CONNECTED) showConnection({ connected: false });
      setStatus(errorMessage(err), "error");
    }
  });
}

// --- Confidential message ("Nachricht") -----------------------------------------------------
// A text instead of a file: the server stores it as a password-protected Vaultwarden Send; the mail gets a card with
// the link, the password goes out like for files.

const messageMax = () => serverConfig?.messageMaxLength ?? 1000;

function renderMessageCount() {
  const length = messageInput.value.length;
  messageCount.textContent = t("message.count", { count: length, max: messageMax() });
  messageCount.classList.toggle("over", length > messageMax());
}
messageInput.addEventListener("input", renderMessageCount);

async function insertMessage() {
  const text = messageInput.value.trim();
  if (!text) {
    messageInput.focus();
    return setStatus(t("message.empty"), "error");
  }
  if (text.length > messageMax()) {
    messageInput.focus();
    return setStatus(t("message.tooLong", { max: messageMax() }), "error");
  }
  const problem = settingsProblem("message");
  if (problem) return setStatus(problem, "error");

  await runInsert(async () => {
    try {
      setStatus(t("message.creating"));
      const result = await api<{ shareUrl: string; expireDate: string; jobId: string }>(
        "/api/messages",
        {
          method: "POST",
          body: JSON.stringify({ ...(await shareSettingsBody(newStep())), text, messageDays: sendDays(messageDaysInput).value }),
        },
        await getToken(),
      );
      // Same password and password mail as the files of this mail (one step).
      await rememberShareChoice(result.jobId);
      const expire = formatDate(result.expireDate);
      await footerReady;
      await insertIntoBody(
        messageCard(result.shareUrl, expire, mailFooter),
        messageText(result.shareUrl, expire, mailFooter),
        shareToken(result.shareUrl),
      );
      messageInput.value = "";
      renderMessageCount();
      setStatus(t("message.inserted"), "success");
    } catch (err) {
      setStatus(errorMessage(err), "error");
    }
  });
}

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
  /** How the password went out after sending (mail or SMS with delivery status); null before / for older links. */
  password?: { channel: "mail" | "sms"; to: string; status: "submitted" | "delivered" | "failed" | "unknown" | null; at: number; deliveredAt: number | null } | null;
}

/** "Kennwort per SMS an +49 170 •••• 78 – zugestellt um 14:05" for the "Meine Freigaben" list. */
function passwordText(p: NonNullable<ShareItem["password"]>): string {
  const time = (ms: number) => new Date(ms).toLocaleTimeString(locale(), { hour: "2-digit", minute: "2-digit" });
  if (p.channel === "mail") return t("shares.passwordMail", { to: p.to });
  const status =
    p.status === "delivered"
      ? p.deliveredAt ? t("shares.smsDeliveredAt", { time: time(p.deliveredAt) }) : t("shares.smsDelivered")
      : p.status === "failed" ? t("shares.smsFailed")
      : t("shares.smsSubmitted");
  return t("shares.passwordSms", { to: p.to, status });
}

const REQUEST_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M10 13V4m0 0L6.5 7.5M10 4l3.5 3.5M4 12v3a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-3" fill="none" ' +
  'stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';

const formatDateTime = (iso: string) =>
  new Date(iso).toLocaleString(locale(), { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });

/** Status line: downloads for files, received uploads for upload links, or "abgelaufen". */
function shareState(share: ShareItem): { text: string; done: boolean; expired: boolean } {
  const expired = share.expiration !== null && share.expiration < new Date().toISOString().slice(0, 10);
  if (expired) return { text: t("shares.expired"), done: false, expired: true };
  if (share.kind === "request") {
    if (share.received === null) return { text: t("shares.statusUnknown"), done: false, expired: false };
    return share.received > 0
      ? { text: t("shares.received", { count: share.received }), done: true, expired: false }
      : { text: t("shares.noneReceived"), done: false, expired: false };
  }
  if (!share.downloads) return { text: t("shares.downloadUnknown"), done: false, expired: false };
  return share.downloads.count > 0
    ? {
        text: share.downloads.last
          ? t("shares.downloadedLast", { count: share.downloads.count, time: formatDateTime(share.downloads.last) })
          : t("shares.downloaded", { count: share.downloads.count }),
        done: true,
        expired: false,
      }
    : { text: t("shares.notDownloaded"), done: false, expired: false };
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
    `<div class="share-meta"></div><div class="share-meta share-password" hidden></div>` +
    `<div class="share-state${state.done ? " done" : ""}${state.expired ? " expired" : ""}">` +
    `<span class="dot-state" aria-hidden="true"></span><span class="state-text"></span></div>` +
    `<div class="share-actions">` +
    // Short labels: the taskpane is narrow (Outlook sets its width); the full text is in title / aria-label.
    `<button type="button" class="btn btn-subtle copy"></button>` +
    `<button type="button" class="btn btn-subtle insert"></button>` +
    `<button type="button" class="btn btn-subtle btn-danger revoke"></button>` +
    `</div></div>`;
  const label = (button: HTMLButtonElement, text: string, full: string) => {
    button.textContent = text;
    button.title = full;
    button.setAttribute("aria-label", full);
  };
  label(row.querySelector<HTMLButtonElement>(".copy")!, t("shares.copy"), t("shares.copyLabel"));
  label(row.querySelector<HTMLButtonElement>(".insert")!, t("shares.insert"), t("shares.insertLabel"));
  label(row.querySelector<HTMLButtonElement>(".revoke")!, t("shares.revoke"), t("shares.revokeLabel"));
  row.querySelector<HTMLElement>(".share-name")!.textContent = share.name;
  row.querySelector<HTMLElement>(".share-name")!.title = share.name;
  row.querySelector<HTMLElement>(".share-meta")!.textContent = [
    t("shares.sharedOn", { date: new Date(share.sharedAt).toLocaleDateString(locale()) }),
    share.expiration ? t("shares.validUntil", { date: formatDate(share.expiration) }) : "",
    share.protected === undefined ? "" : t(share.protected ? "shares.withPassword" : "shares.withoutPassword"),
  ]
    .filter(Boolean)
    .join(" · ");
  row.querySelector<HTMLElement>(".state-text")!.textContent = state.text;
  if (share.password) {
    const line = row.querySelector<HTMLElement>(".share-password")!;
    line.textContent = passwordText(share.password);
    line.hidden = false;
    line.classList.toggle("failed", share.password.status === "failed");
  }

  const copy = row.querySelector<HTMLButtonElement>(".copy")!;
  copy.addEventListener("click", async () => {
    await copyText(share.url);
    copy.textContent = t("shares.copied");
    setTimeout(() => (copy.textContent = t("shares.copy")), 2000);
  });

  // A new link for this mail with the settings of the "Senden" tab – the original password cannot be recovered.
  const insert = row.querySelector<HTMLButtonElement>(".insert")!;
  insert.hidden = readMode; // a mail being read cannot take a link
  insert.addEventListener("click", () => {
    linkAgain(share, insert).catch(() => undefined);
  });

  // Two-step confirmation instead of a modal dialog (add-ins cannot show native dialogs).
  const revoke = row.querySelector<HTMLButtonElement>(".revoke")!;
  let armedUntil = 0;
  revoke.addEventListener("click", async () => {
    if (Date.now() > armedUntil) {
      armedUntil = Date.now() + 4000;
      revoke.textContent = t("shares.confirm");
      revoke.setAttribute("aria-label", t("shares.confirmLabel"));
      revoke.classList.add("confirm");
      setTimeout(() => {
        if (Date.now() >= armedUntil) {
          revoke.textContent = t("shares.revoke");
          revoke.setAttribute("aria-label", t("shares.revokeLabel"));
          revoke.classList.remove("confirm");
        }
      }, 4100);
      return;
    }
    revoke.disabled = true;
    try {
      const result = await api<{ file?: "deleted" | "scheduled" | "kept"; deleteAfterDays?: number | null }>(
        `/api/shares/${encodeURIComponent(share.id)}`,
        { method: "DELETE" },
        await getToken(),
      );
      row.remove();
      // A revoked link counts as expired: the server moved the file to the trash or scheduled it (CLEANUP_AFTER_DAYS).
      const fate =
        share.kind === "request"
          ? t("shares.fateRequest")
          : result.file === "deleted"
            ? t("shares.fateDeleted")
            : result.file === "scheduled"
              ? t("shares.fateScheduled", { count: result.deleteAfterDays ?? 0 })
              : t("shares.fateKept");
      setStatus(t("shares.revoked", { name: share.name, fate }), "success");
      renderEmptyLists();
    } catch (err) {
      revoke.disabled = false;
      setStatus(t("shares.revokeFailed", { error: errorMessage(err) }), "error");
    }
  });
  return row;
}

interface LinkAgainResult {
  kind: "file" | "request";
  shareUrl: string;
  expireDate: string;
  jobId: string | null;
  noPassword?: boolean;
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
    setStatus(t("shares.settingsProblem", { problem }), "error");
    return;
  }
  button.disabled = true;
  busy++;
  try {
    setStatus(t("shares.linking", { name: share.name }));
    const result = await api<LinkAgainResult>(
      `/api/shares/${encodeURIComponent(share.id)}/link`,
      { method: "POST", body: JSON.stringify(await shareSettingsBody(newStep())) },
      await getToken(),
    );
    await rememberShareChoice(result.jobId, isWithoutPassword(result));
    const expire = formatDate(result.expireDate);
    const withPassword = !isWithoutPassword(result);
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
      sentences(
        t("shares.linked", { name: share.name }),
        t(withPassword ? "submit.passwordFollows" : "shares.linkedNoPassword"),
      ),
      "success",
    );
    loadShares().catch(() => undefined); // the new link shows up in the list
  } catch (err) {
    if (err instanceof ApiError && err.code === NOT_CONNECTED) showConnection({ connected: false });
    setStatus(t("insert.failed", { error: errorMessage(err) }), "error");
  } finally {
    button.disabled = false;
    nextStepPassword();
    busy--;
  }
}

async function loadShares() {
  refreshSharesButton.disabled = true;
  try {
    const shares = await api<ShareItem[]>("/api/shares", {}, await getToken());
    if (readMode) {
      // The links of the mail being read on top, the others below.
      const tokens = await mailTokens();
      const inMail = (share: ShareItem) => tokens.has(shareToken(share.url));
      mailShares.replaceChildren(...shares.filter(inMail).map(shareRow));
      sharesList.replaceChildren(...shares.filter((share) => !inMail(share)).map(shareRow));
    } else {
      sharesList.replaceChildren(...shares.map(shareRow));
    }
    renderEmptyLists();
  } finally {
    refreshSharesButton.disabled = false;
  }
}

/** "Noch keine Freigaben" / "keine Links in dieser Nachricht" when a list is (or became) empty. */
function renderEmptyLists() {
  noShares.hidden = sharesList.childElementCount > 0 || mailShares.childElementCount > 0;
  mailNone.hidden = mailShares.childElementCount > 0;
}

/**
 * Share tokens of the links in the mail being read – Nextcloud links end in /s/<token>. Only the user's own shares
 * are matched against them, so other links in the mail do no harm.
 */
async function mailTokens(): Promise<Set<string>> {
  const item = Office.context.mailbox.item as Office.MessageRead | null | undefined;
  if (!item) return new Set(); // pinned pane, no mail selected
  const html = await officeAsync<string>((cb) => item.body.getAsync(Office.CoercionType.Html, cb)).catch(() => "");
  return new Set(Array.from(html.matchAll(/\/s\/([A-Za-z0-9]+)/g), (match) => match[1]));
}

/** Read mode: no "Senden" tab, only the shares; follows the selected mail while the pane is pinned. */
async function startReadView() {
  tabList.hidden = true;
  panelSend.hidden = true;
  panelShares.hidden = false;
  mailSection.hidden = false;
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      loadShares().catch((err) => setStatus(errorMessage(err), "error"));
    });
  } catch {
    // Pinning is not available in every Outlook version – the pane then closes with the mail anyway.
  }
  try {
    await refreshConnection(); // loads the shares once connected (showConnection)
  } catch (err) {
    setStatus(t("signInFailed", { error: errorMessage(err) }), "error");
  }
}

refreshSharesButton.addEventListener("click", () => {
  loadShares().catch((err) => setStatus(errorMessage(err), "error"));
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
    // Also for the send handler: it starts with this cached copy (attachment policy, diagnostics switch).
    try {
      Office.context.roamingSettings.set(CONFIG_CACHE_KEY, serverConfig);
      Office.context.roamingSettings.saveAsync();
    } catch {
      // Caching is best effort.
    }
  } catch {
    setStatus(t("unreachable", { url: API_BASE }), "error");
    return;
  }
  if (readMode) return startReadView();
  footerReady = resolveFooter().catch(() => undefined); // in the background – needed when a link is inserted

  const state = await readDraftState();
  expireInput.value = String(serverConfig.defaultExpireDays);
  expireInput.max = String(serverConfig.maxExpireDays);
  maxAccessInput.max = String(serverConfig.sendMaxAccessLimit);
  for (const input of [passwordDaysInput, messageDaysInput]) {
    input.max = String(sendMaxDays());
    input.value = String(sendDefaultDays());
  }
  maxAccessInput.value = String(state.maxAccess ?? serverConfig.sendMaxAccessCount);
  if (state.password) {
    passwordInput.value = state.password;
    lockPassword();
  } else if (state.noPassword) {
    lockNoPassword();
  } else if (state.separate) {
    lockSeparate();
    setPasswordMode("random");
  } else {
    setPasswordMode("random");
  }
  // "Kennwort senden an": SMS only when the server has it configured; restore the choice of this draft.
  if (!serverConfig.smsEnabled) channelSmsOption.remove();
  renderMessageCount();
  passwordToInput.value = state.passwordTo ?? "";
  passwordSmsInput.value = state.passwordSms ?? "";
  showChannel(state.passwordSms && serverConfig.smsEnabled ? "sms" : state.passwordTo ? "mail" : "recipients");
  renderKind();
  // Outlook has no "body changed" event: check when the pane gets focus and every few seconds while locked.
  await releaseIfLinksDeleted();
  window.addEventListener("focus", () => void releaseIfLinksDeleted());
  setInterval(() => void releaseIfLinksDeleted(), 4000);

  await renderAttachments().catch(() => undefined);
  composeItem().addHandlerAsync(Office.EventType.AttachmentsChanged, () => {
    renderAttachments().catch(() => undefined);
  });

  try {
    await refreshConnection();
  } catch (err) {
    setStatus(t("signInFailed", { error: errorMessage(err) }), "error");
  }
});
