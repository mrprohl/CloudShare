import {
  api,
  attachmentThreshold,
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
import { inLanguage, languageChoices, locale, t } from "../shared/i18n";
import { errorText } from "../shared/sso";
import { filesCard, filesText, messageCard, messageText, requestCard, requestText, type CardProof } from "./mailCard";
import { initSettings, loadSettings, userSettings } from "./settings";
import { labelCatalog, mailLabel, readFileLabel, setMailLabel, type FileLabel, type LabelDetails } from "./labels";

// All visible texts come from locales/<language>.json (t(), see shared/i18n.ts) and follow Outlook's terminology
// ("Nachricht"/"message", "Kennwort"/"password").

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

const connectView = $<HTMLElement>("connect-view");
const connectButton = $<HTMLButtonElement>("connect");
const connectHint = $<HTMLParagraphElement>("connect-hint");
const connectLink = $<HTMLAnchorElement>("connect-link");
const sendView = $<HTMLElement>("send-view");
const accountText = $<HTMLElement>("account");
const accountToggle = $<HTMLButtonElement>("account-toggle");
const accountMenu = $<HTMLDivElement>("account-menu");
const accountStorage = $<HTMLParagraphElement>("account-storage");
const accountMenuName = $<HTMLParagraphElement>("account-menu-name");
const storageMeter = $<HTMLDivElement>("storage-meter");
const storageMeterFill = $<HTMLDivElement>("storage-meter-fill");
const disconnectButton = $<HTMLButtonElement>("disconnect");
const disconnectText = $<HTMLSpanElement>("disconnect-text");
const nextcloudOpen = $<HTMLButtonElement>("nextcloud-open");
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
const modeSame = $<HTMLInputElement>("mode-same");
const modeSameRow = $<HTMLLabelElement>("mode-same-row");
const passwordGroup = $<HTMLDivElement>("password-group");
const maxAccessField = $<HTMLDivElement>("max-access-field");
const passwordDaysField = $<HTMLDivElement>("password-days-field");
const passwordDaysInput = $<HTMLInputElement>("password-days");
const messageDaysInput = $<HTMLInputElement>("message-days");
const passwordToField = $<HTMLDivElement>("password-to-field");
const recipientLanguageField = $<HTMLDivElement>("recipient-language-field");
const recipientLanguageSelect = $<HTMLSelectElement>("recipient-language");
const proofField = $<HTMLDivElement>("proof-field");
const proofInput = $<HTMLInputElement>("proof");
const encryptedField = $<HTMLDivElement>("encrypted-field");
const encryptedText = $<HTMLParagraphElement>("encrypted-text");
const encryptedConfirm = $<HTMLInputElement>("encrypted-confirm");
const encryptedConfirmRow = $<HTMLLabelElement>("encrypted-confirm-row");
/** Names of the encrypted files the note was shown for ("" = none). */
let encryptedShown = "";
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
const expireHint = $<HTMLParagraphElement>("expire-hint");
const tabSend = $<HTMLButtonElement>("tab-send");
const tabShares = $<HTMLButtonElement>("tab-shares");
const tabMail = $<HTMLButtonElement>("tab-mail");
const refreshSharesButton = $<HTMLButtonElement>("refresh-shares");
const panelMail = $<HTMLDivElement>("panel-mail");
const helpToggle = $<HTMLButtonElement>("help-toggle");
const helpClose = $<HTMLButtonElement>("help-close");
const panelSend = $<HTMLDivElement>("panel-send");
const panelShares = $<HTMLDivElement>("panel-shares");
const panelHelp = $<HTMLElement>("panel-help");
const panelSettings = $<HTMLElement>("panel-settings");
const settingsClose = $<HTMLButtonElement>("settings-close");
/** "Einstellungen" in the account menu. */
const settingsOpen = $<HTMLButtonElement>("settings-open");
const panelProofs = $<HTMLElement>("panel-proofs");
const proofsOpen = $<HTMLButtonElement>("proofs-open");
const proofsClose = $<HTMLButtonElement>("proofs-close");
const proofsSearch = $<HTMLInputElement>("proofs-search");
const proofsList = $<HTMLDivElement>("proofs-list");
const proofsEmpty = $<HTMLParagraphElement>("proofs-empty");
const mailShares = $<HTMLDivElement>("mail-shares");
const mailNone = $<HTMLParagraphElement>("mail-none");
const sharesList = $<HTMLDivElement>("shares");
const noShares = $<HTMLParagraphElement>("no-shares");
const expireInput = $<HTMLInputElement>("expire");
const maxAccessInput = $<HTMLInputElement>("max-access");
const submitButton = $<HTMLButtonElement>("submit");
const progress = $<HTMLDivElement>("progress");
const progressBar = $<HTMLDivElement>("progress-bar");
const statusBar = $<HTMLDivElement>("status");

/**
 * Opened from a mail being read (manifest: taskpane.html?view=read, e.g. in "Sent Items"): only "Verlauf",
 * with the links found in that mail on top. Nothing is inserted or changed in the mail.
 */
const readMode = new URLSearchParams(location.search).get("view") === "read";

let serverConfig: ServerConfig;
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

/** "progress": grey like "info", but stays until the next message (uploading, creating …) – no timer. */
type StatusKind = "info" | "progress" | "success" | "error";

/** The sign inside the circle of each kind of message. */
const STATUS_GLYPHS: Record<StatusKind, string> = {
  info: '<path d="M8 7v4M8 4.8v.4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>',
  progress: '<path d="M8 7v4M8 4.8v.4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>',
  success: '<path d="m5.3 8.2 1.8 1.8 3.6-3.8" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>',
  error: '<path d="M8 4.5v4.2M8 11v.4" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"/>',
};
const STATUS_CIRCLE = '<circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor"/>';

/**
 * A message that runs down (SUCCESS_MESSAGE_SECONDS, like the OTP timer in Vaultwarden): a faint track and the ring
 * around the tick (or the "!"), emptying clockwise from 12 o'clock. It stops while the pointer (or the keyboard focus)
 * is on the message – taskpane.css – and the message goes when it has run down.
 */
const statusTimer = (kind: StatusKind) =>
  '<circle cx="8" cy="8" r="6.5" fill="none" stroke="currentColor" opacity=".25"/>' +
  '<circle class="status-timer" cx="8" cy="8" r="6.5" fill="none" stroke="currentColor" pathLength="100" stroke-dasharray="100"/>' +
  STATUS_GLYPHS[kind];

/**
 * Seconds a message stays (server setting, 5 until the configuration is loaded); 0: it stays.
 * Progress ("progress") and a message with a button ("Erneut versuchen") stay until the next one; grey notes ("info")
 * run down like the others.
 */
const messageSeconds = () => serverConfig?.successMessageSeconds ?? 5;

/** A button in the message, e.g. "Erneut versuchen" after a failed start. */
type StatusAction = { label: string; run: () => void };

/** "Erneut versuchen": CloudShare starts afresh, as if the pane was opened again (sign-in, server, settings). */
const retryAction = (): StatusAction => ({ label: t("retry"), run: () => location.reload() });

let statusShown = 0;
function setStatus(text: string, kind: StatusKind = "info", action?: StatusAction | StatusAction[]) {
  const shown = ++statusShown;
  const actions = action ? (Array.isArray(action) ? action : [action]) : [];
  statusBar.hidden = !text;
  statusBar.className = `messagebar ${kind}`;
  // No timer for an empty message (setStatus("") only clears the bar), progress and a message with a button.
  const seconds = text && kind !== "progress" && actions.length === 0 ? messageSeconds() : 0;
  statusBar.innerHTML = text
    ? `<svg class="icon" viewBox="0 0 16 16" aria-hidden="true">${seconds > 0 ? statusTimer(kind) : STATUS_CIRCLE + STATUS_GLYPHS[kind]}</svg><span></span>`
    : "";
  if (text) statusBar.querySelector("span")!.textContent = text;
  for (const item of text ? actions : []) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "link-button status-action";
    button.textContent = item.label;
    button.addEventListener("click", item.run);
    statusBar.querySelector("span")!.append(" ", button);
  }
  if (seconds > 0) {
    const ring = statusBar.querySelector<SVGCircleElement>(".status-timer")!;
    ring.style.animationDuration = `${seconds}s`;
    ring.addEventListener("animationend", () => {
      if (shown === statusShown) setStatus(""); // unless another message took its place meanwhile
    });
  }
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

/** Storage of the account (account menu): "37,5 GB von 50 GB belegt" (quota), "12,3 GB frei" (free space only), "3,2 GB belegt". */
function storageText(status: ConnectionStatus): string {
  const isBytes = (value: unknown): value is number => typeof value === "number" && value >= 0;
  if (isBytes(status.freeBytes) && isBytes(status.totalBytes)) {
    // As OneDrive and Outlook say it, and as the bar shows it: "37,5 GB von 50 GB belegt".
    const used = isBytes(status.usedBytes) ? status.usedBytes : Math.max(0, status.totalBytes - status.freeBytes);
    return t("storage.usedOf", { used: formatSize(used), total: formatSize(status.totalBytes) });
  }
  if (isBytes(status.freeBytes)) return t("storage.free", { free: formatSize(status.freeBytes) });
  if (isBytes(status.usedBytes)) return t("storage.used", { used: formatSize(status.usedBytes) });
  return "";
}

function showConnection(status: ConnectionStatus) {
  connectView.hidden = status.connected;
  sendView.hidden = !status.connected;
  connectHint.hidden = true;
  // The bottom bar shows the Nextcloud account; its menu the whole name, the storage and the entries.
  accountText.textContent = status.loginName ?? "";
  // The bar may cut a long name short: hovering shows it whole.
  // That it opens a menu shows the arrow; screen readers hear what it holds (aria-describedby).
  accountToggle.title = status.loginName ? t("account.signedIn", { name: status.loginName }) : "";
  const storage = storageText(status);
  // The menu head: the whole name, the storage as a bar (with a quota) and as text.
  accountMenuName.textContent = status.loginName ?? "";
  accountStorage.textContent = storage;
  accountStorage.hidden = !storage;
  const isBytes = (value: unknown): value is number => typeof value === "number" && value >= 0;
  const share =
    isBytes(status.freeBytes) && isBytes(status.totalBytes) && status.totalBytes > 0
      ? Math.min(1, Math.max(0, 1 - status.freeBytes / status.totalBytes))
      : null;
  storageMeter.hidden = share === null;
  storageMeterFill.style.width = `${Math.round((share ?? 0) * 100)}%`;
  storageMeter.classList.toggle("full", (share ?? 0) >= 0.9);
  // "Nextcloud öffnen": the address of the Nextcloud CloudShare uses (from /api/config).
  nextcloudOpen.hidden = !serverConfig?.nextcloudUrl;
  nextcloudOpen.title = serverConfig?.nextcloudUrl ?? "";
  showAccountMenu(false);
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

// "Nextcloud öffnen" in the account menu: the Nextcloud in the default browser – its own session there, or a sign-in
// (single sign-on). The app password of the connection never leaves the server.
nextcloudOpen.addEventListener("click", () => {
  const url = serverConfig?.nextcloudUrl;
  if (!url) return;
  showAccountMenu(false);
  const ui = Office.context.ui as unknown as { openBrowserWindow?: (url: string) => void };
  if (ui.openBrowserWindow && Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
    ui.openBrowserWindow(url);
  } else {
    window.open(url, "_blank", "noopener");
  }
});

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

async function disconnect(e: Event) {
  e.preventDefault();
  try {
    await api("/api/nextcloud/connection", { method: "DELETE" }, await getToken());
    showConnection({ connected: false });
    setStatus(t("disconnect.done"));
  } catch (err) {
    setStatus(t("disconnect.failed", { error: errorMessage(err) }), "error");
  }
}
// --- Account menu (bottom bar, like the profile menu in Outlook) ----------------------

/** Opens or closes the account menu above the bar: "Einstellungen", "Hilfe", "Verbindung trennen". */
function showAccountMenu(open: boolean) {
  accountMenu.hidden = !open;
  accountToggle.setAttribute("aria-expanded", String(open));
  if (!open) armDisconnect(false);
}

/**
 * "Verbindung trennen" takes two clicks, like "Zurückziehen": the first one turns it into "Wirklich trennen?" (red)
 * for 4 s, the second one disconnects. Closing the menu disarms it.
 */
let disconnectArmedUntil = 0;
function armDisconnect(armed: boolean) {
  disconnectArmedUntil = armed ? Date.now() + 4000 : 0;
  disconnectText.textContent = t(armed ? "disconnect.confirm" : "disconnect.link");
  disconnectButton.classList.toggle("confirm", armed);
}
accountToggle.addEventListener("click", () => {
  const open = accountMenu.hidden === true;
  showAccountMenu(open);
  if (open) accountMenu.querySelector<HTMLButtonElement>("button:not([hidden])")?.focus(); // the first entry
});
// Closes with Escape (focus back on the account) or a click anywhere else.
accountMenu.addEventListener("keydown", (event) => {
  if (event.key !== "Escape") return;
  showAccountMenu(false);
  accountToggle.focus();
});
document.addEventListener("click", (event) => {
  if (!accountMenu.hidden && !accountMenu.contains(event.target as Node) && !accountToggle.contains(event.target as Node)) {
    showAccountMenu(false);
  }
});
disconnectButton.addEventListener("click", (event) => {
  if (Date.now() > disconnectArmedUntil) {
    armDisconnect(true);
    setTimeout(() => {
      if (disconnectArmedUntil && Date.now() >= disconnectArmedUntil) armDisconnect(false);
    }, 4100);
    return;
  }
  showAccountMenu(false);
  void disconnect(event);
});

// --- File lists ------------------------------------------------------------------

const FILE_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M5.5 2.5h5.3l3.7 3.7v10.3a1 1 0 0 1-1 1h-8a1 1 0 0 1-1-1v-13a1 1 0 0 1 1-1Z" fill="none" stroke="currentColor"/>' +
  '<path d="M10.5 2.5v4h4" fill="none" stroke="currentColor"/></svg>';

/**
 * One row in Outlook's list style: checkbox, file icon, name and an optional "Groß" badge; below the name the file's
 * sensitivity label once it is known (showLabel) – no size.
 */
function fileRow(
  name: string,
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
    `<span class="file-text"><span class="file-name"></span><span class="file-meta" hidden></span></span>` +
    (large ? `<span class="badge"></span>` : "");
  const box = row.querySelector("input")!;
  box.value = value;
  box.checked = checked;
  box.dataset.name = name;
  row.querySelector(".file-name")!.textContent = name;
  row.querySelector<HTMLElement>(".file-name")!.title = name;
  const badge = row.querySelector<HTMLElement>(".badge");
  if (badge) badge.textContent = t("files.large");
  // Own name for the checkbox: otherwise a remove button inside the label would become part of it.
  box.setAttribute("aria-label", name);
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

// --- Sensitivity labels (Microsoft Purview) -----------------------------------------------------------
// labels.ts reads the label from the file; here it shows in the list, a protected file needs a confirmation, and the
// mail gets the highest label of its files (LABEL_POLICY: inherit / suggest / off, LABEL_ENCRYPTED: warn / block / allow).

const labelPolicy = () => serverConfig?.labelPolicy ?? "inherit";
const encryptedPolicy = () => serverConfig?.labelEncrypted ?? "warn";
/** Attachments up to this size are read for the list right away; larger ones when they are inserted. */
const LABEL_PREVIEW_MAX = 50 * MB;

let catalogPromise: Promise<LabelDetails[] | null> | undefined;
const catalog = () => (catalogPromise ??= labelCatalog());
/** Label per attachment ID or picked file – each file is read once. */
const labelCache = new Map<string | File, Promise<FileLabel | null>>();
function labelOf(key: string | File, name: string, content: () => Promise<Blob>): Promise<FileLabel | null> {
  let label = labelCache.get(key);
  if (!label) {
    label = content().then((blob) => readFileLabel(blob, name)).catch(() => null);
    labelCache.set(key, label);
  }
  return label;
}

/** The label's name from Outlook's catalog; without it (other organisation, no catalog) the general wording. */
async function labelName(label: FileLabel): Promise<{ name: string; known?: LabelDetails }> {
  const known = label.id ? (await catalog())?.find((l) => l.id === label.id) : undefined;
  if (known) return { name: known.name, known };
  return { name: t(label.id ? "labels.unknownName" : "labels.protected") };
}

const LOCK_ICON =
  '<svg class="icon label-lock" viewBox="0 0 16 16" aria-hidden="true"><rect x="3.5" y="7" width="9" height="6.5" rx="1" ' +
  'fill="none" stroke="currentColor"/><path d="M5.5 7V5a2.5 2.5 0 0 1 5 0v2" fill="none" stroke="currentColor"/></svg>';

/** "Vertraulich" (a lock for an encrypted file) under the name of a row, once the label is known. */
async function showLabel(row: HTMLElement, label: Promise<FileLabel | null>) {
  const found = await label;
  if (!found || !row.isConnected) return;
  const { name } = await labelName(found);
  const span = document.createElement("span");
  span.className = "file-label";
  span.innerHTML = found.encrypted ? LOCK_ICON : "";
  span.append(name);
  const meta = row.querySelector<HTMLElement>(".file-meta")!;
  meta.append(span);
  meta.hidden = false;
  const box = row.querySelector("input")!;
  box.setAttribute("aria-label", `${box.getAttribute("aria-label")}, ${found.encrypted ? t("labels.encryptedLabel", { name }) : name}`);
  if (found.encrypted) void renderEncrypted();
}

/** Ticked files (attachments and files from the computer) – the button says how many get a link. */
const selectedFileCount = () =>
  attachmentList.querySelectorAll("input:checked").length + pickedList.querySelectorAll("input:checked").length;

/** The size (MB) the list was last rendered with – a change of recipients re-renders only if it changes. */
let shownThreshold: number | undefined;

async function renderAttachments() {
  const attachments = await getFileAttachments();
  // The size that applies to this draft – outside or inside the organisation (attachmentThreshold).
  shownThreshold = await attachmentThreshold(serverConfig);
  const threshold = shownThreshold * MB;
  // Large attachments are preselected.
  // ATTACHMENT_THRESHOLD_MB=0: every attachment is preselected, and none is called "Groß".
  const rows = attachments.map((a) => fileRow(a.name, a.id, a.size > threshold, a.size > threshold && threshold > 0));
  attachmentList.replaceChildren(...rows);
  // Their sensitivity label under the name (larger attachments are read only when they are inserted).
  if (labelPolicy() !== "off") {
    attachments.forEach((a, i) => {
      if (a.size <= LABEL_PREVIEW_MAX) void showLabel(rows[i], labelOf(a.id, a.name, () => readAttachment(a.id)));
    });
  }
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
      fileRow(file.name, String(i), !unticked.has(file), false, () => {
        pickedFiles = pickedFiles.filter((f) => f !== file);
        renderPicked();
      }),
    ),
  );
  // Only a few small parts of a file are read for its label, so every size is fine.
  if (labelPolicy() !== "off") {
    pickedFiles.forEach((file, i) => void showLabel(pickedList.children[i] as HTMLElement, labelOf(file, file.name, async () => file)));
  }
  noAttachments.hidden = attachmentList.childElementCount > 0 || pickedFiles.length > 0;
  renderSubmit();
}

/** The files the rows of #picked currently show (their checkbox values are indexes into it). */
let pickedRendered: File[] = [];

// Ticking another file may bring an encrypted one in or out (the note below the list).
for (const list of [attachmentList, pickedList]) list.addEventListener("change", () => void renderEncrypted());


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
  /** "Zustellnachweis": tracking ID of the mail's case (null: not ticked). */
  proofId?: string | null;
}

/** "Zustellnachweis" in the card: the tracking ID; `gateway` – the recipient confirms their address first. */
const cardProof = (proofId: string | null | undefined, gateway: boolean): CardProof | undefined =>
  proofId ? { id: proofId, gateway } : undefined;

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
      setStatus(t("files.uploading", { name: fileName, percent }), "progress");
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
  await rememberShareChoice(result.jobId, isWithoutPassword(result), result.proofId);
  return result;
}

/**
 * Settings of the "Senden" tab for a new link. `step` identifies the click ("Als Link einfügen" etc.): every step
 * gets its own password and password mail, files of the same step share them (see attachToJob on the server).
 */
/** The draft's subject the server knows (sent with the last link or by syncDraftSubject). */
let syncedSubject: string | undefined;

/** Subject of the draft (undefined if Outlook cannot tell) – "Verlauf" shows it until the mail is sent. */
async function draftSubject(): Promise<string | undefined> {
  try {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    return await officeAsync<string>((cb) => item.subject.getAsync(cb));
  } catch {
    return undefined;
  }
}

/**
 * The subject typed after a link was inserted reaches "Verlauf" too: checked with the drafts' other state every few
 * seconds while the pane is open; a change goes to the server (PATCH /api/jobs/:id) – also kept if the draft is then
 * discarded. Only for a mail with links (a job).
 */
async function syncDraftSubject() {
  if (busy > 0 || document.hidden || readMode) return;
  const state = await readDraftState();
  if (!state.jobId) return;
  const subject = await draftSubject();
  if (subject === undefined || subject === syncedSubject) return;
  syncedSubject = subject;
  await api(`/api/jobs/${encodeURIComponent(state.jobId)}`, { method: "PATCH", body: JSON.stringify({ subject }) }, await getToken()).catch(
    () => {
      syncedSubject = undefined; // try again on the next check
    },
  );
}

async function shareSettingsBody(step: string) {
  const expireDays = Number(expireInput.value);
  const subject = await draftSubject();
  syncedSubject = subject;
  if (withoutPassword()) {
    return { noPassword: true, expireDays, jobId: (await readDraftState()).jobId, language: recipientLanguage(), subject, ...(await proofRequest()) };
  }
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
    // "Kennwort senden an" of this link (per insertion): links of one step with different ways get the password on
    // every way – e.g. one password mail to the recipients and one SMS.
    ...(wanted?.passwordSms ? { passwordSms: wanted.passwordSms } : { passwordTo: wanted?.passwordTo ?? "" }),
    // Language of this step's password mail and SMS ("Sprache für die Empfänger").
    language: recipientLanguage(),
    subject,
    // "Zustellnachweis" for this link (only when ticked).
    ...(await proofRequest()),
  };
}

/**
 * "Sprache für die Empfänger": the languages CloudShare has (by their own name), `selected` preselected – this mail's
 * choice or the one from the settings. Hidden while there is only one language.
 */
function renderRecipientLanguages(selected: string) {
  const choices = languageChoices();
  recipientLanguageSelect.replaceChildren(...choices.map((l) => new Option(l.name, l.code)));
  recipientLanguageSelect.value = choices.some((l) => l.code === selected) ? selected : (choices[0]?.code ?? "");
  recipientLanguageField.hidden = choices.length < 2;
  renderValidity(); // the summary names the language
}

/** The language the recipients get the link card, password mail and SMS in. */
const recipientLanguage = () => recipientLanguageSelect.value || undefined;

/** "Zustellnachweis" ticked (and offered by the server). */
function proofWanted() {
  return !proofField.hidden && proofInput.checked;
}

/**
 * What a link request says about "Zustellnachweis": ticked, the case of this mail so far (its tracking ID) and the
 * client – Outlook's host, platform and version for the record.
 */
async function proofRequest() {
  if (!proofWanted()) return {};
  const diagnostics = Office.context.diagnostics;
  return {
    proof: true,
    proofId: (await readDraftState()).proofId,
    client: { host: String(diagnostics?.host ?? ""), platform: String(diagnostics?.platform ?? ""), version: diagnostics?.version ?? "" },
  };
}

proofInput.addEventListener("change", () => {
  void updateDraftState({ proof: proofInput.checked });
  renderValidity();
});

// A choice for this mail is kept with the draft: the pane opened again shows it, and new links follow it.
recipientLanguageSelect.addEventListener("change", () => {
  renderValidity();
  void updateDraftState({ recipientLanguage: recipientLanguageSelect.value });
});

/** New ID for one click on "Als Link einfügen" / "Link zum Hochladen einfügen" / "Verlauf → Als Link einfügen". */
const newStep = () => Math.random().toString(36).slice(2, 12);

/**
 * From the first link on, the choice (this password or none) applies to the whole mail. `proofId`: the link joined the
 * mail's "Zustellnachweis" – later links with the box ticked join the same case.
 */
async function rememberShareChoice(jobId: string | null, noPassword = jobId === null, proofId?: string | null) {
  if (proofId) await updateDraftState({ proofId });
  if (noPassword || jobId === null) {
    // Links without password. With the download confirmation the server keeps a job without password, which the
    // send handler arms like a password job (it only waits for the sent mail).
    await updateDraftState(jobId ? { noPassword: true, jobId } : { noPassword: true });
    if (jobId) await tagDraft(jobId);
    lockNoPassword();
    return;
  }
  // The way of "Kennwort senden an" travels with every link (shareSettingsBody) – nothing to hand to the job here.
  if ((await readDraftState()).password) {
    // Mail started before passwords per step existed: it keeps its one shared password.
    await updateDraftState({ jobId });
    await tagDraft(jobId);
  } else {
    // One password per step; this step is the one "Gleiches Kennwort wie zuvor" joins next (nextStepPassword).
    if (pendingStep) lastStep = pendingStep;
    await updateDraftState({ jobId, separate: true, ...(lastStep ? { lastStep } : {}) });
    await tagDraft(jobId);
    lockSeparate();
  }
}

// --- "Kennwort senden an": recipients of the mail / another mail address / SMS ------------------------------

const looksLikeEmail = (value: string) => /^[^\s@,;<>"()]+@[^\s@,;<>"()]+\.[^\s@,;<>"()]{2,}$/.test(value);

/**
 * Same rules as the server (server/src/sms.ts): "tel:" and the "(0)" of "+49 (0)170 …" removed, E.164, "00" → "+",
 * leading "0" → default country code.
 */
function normalizePhone(input: string): string | null {
  let number = input
    .trim()
    .replace(/^tel:/i, "")
    .replace(/＋/g, "+")
    .replace(/(\d)\s*\(0\)/g, "$1")
    .replace(/[\s\-‐‑–—./()]/g, "");
  if (number.startsWith("00")) number = `+${number.slice(2)}`;
  else if (/^0[1-9]/.test(number)) number = `+${serverConfig.smsDefaultCountry ?? "49"}${number.slice(1)}`;
  return /^\+[1-9]\d{7,14}$/.test(number) ? number : null;
}

/** SMS_ALLOWED_NUMBERS / SMS_BLOCKED_NUMBERS from the server; null if empty or not readable here (the server checks). */
function numberRule(source: string | undefined): RegExp | null {
  if (!source) return null;
  try {
    return new RegExp(source);
  } catch {
    return null;
  }
}

/**
 * Why no SMS may go to this number (E.164), or null – the server's checks (smsNumberProblem in server/src/sms.ts):
 * SMS_ALLOWED_COUNTRIES (calling codes are prefix-free, so "starts with +<code>" is exact), then the number rules.
 */
function smsNumberProblem(number: string, scope: "channel" | "resend" = "channel"): string | null {
  const codes = serverConfig.smsCountryCodes ?? [];
  if (codes.length > 0 && !codes.some((code) => number.startsWith(`+${code}`))) {
    return t(`${scope}.smsCountry`, { countries: serverConfig.smsCountries ?? "" });
  }
  const allowed = numberRule(serverConfig.smsAllowedNumbers);
  if (allowed && !allowed.test(number)) return t(`${scope}.smsNumberNotAllowed`);
  if (numberRule(serverConfig.smsBlockedNumbers)?.test(number)) return t(`${scope}.smsNumberBlocked`);
  return null;
}

/** "Phone numbers from Germany and Austria are possible." – only if SMS_ALLOWED_COUNTRIES limits them. */
const smsCountriesHint = () =>
  (serverConfig.smsCountryCodes ?? []).length > 0 ? ` ${t("channel.smsCountries", { countries: serverConfig.smsCountries ?? "" })}` : "";

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
  renderSubmit(); // another address or SMS needs its input first
}

/** "Kennwort senden an" offers the ways chosen in the settings – and always the one this mail already uses. */
function applyChannelChoices(keep: Channel) {
  const offered = new Set<string>(userSettings().channels);
  for (const option of Array.from(channelSelect.options)) {
    const show = offered.has(option.value) || option.value === keep;
    option.hidden = !show;
    option.disabled = !show;
  }
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
    return number && !smsNumberProblem(number) ? { passwordTo: undefined, passwordSms: number } : null;
  }
  return { passwordTo: undefined, passwordSms: undefined };
}

/** Stores the channel for this mail and, once the job exists, on the server. */
async function applyChannel(target: { passwordTo?: string; passwordSms?: string }) {
  const state = await updateDraftState(target);
  try {
    // Per insertion: the choice travels with the next link (shareSettingsBody); links already in the mail keep their
    // way. Only a mail with one shared password from before this version changes its one way on the job.
    if (state.jobId && state.password) {
      await api(`/api/jobs/${encodeURIComponent(state.jobId)}`, {
        method: "PATCH",
        body: JSON.stringify(target.passwordSms ? { passwordSms: target.passwordSms } : { passwordTo: target.passwordTo ?? "" }),
      }, await getToken());
    }
    const saved = target.passwordSms
      ? t("channel.savedSms", { to: target.passwordSms })
      : target.passwordTo
        ? t("channel.savedMail", { to: target.passwordTo })
        : t("channel.savedRecipients");
    setStatus(state.jobId && !state.password ? sentences(saved, t("channel.nextLinkOnly")) : saved, "success");
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
  passwordSmsInput.value = number;
  // Told right away – not only when the link is inserted.
  const numberProblem = smsNumberProblem(number);
  if (numberProblem) {
    passwordSmsInput.setAttribute("aria-invalid", "true");
    setStatus(numberProblem, "error");
    return;
  }
  passwordSmsInput.removeAttribute("aria-invalid");
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
    if (sms) {
      const number = normalizePhone(input.value);
      if (number) return smsNumberProblem(number) ?? t("channel.smsInvalidShort");
      return t(input.value.trim() ? "channel.smsInvalidShort" : "channel.smsMissing");
    }
    return t(input.value.trim() ? "channel.mailInvalidShort" : "channel.mailMissing");
  }
  return null;
}

/** Every date in the pane and the cards: day and month with two digits and the full year ("02.12.2026"). */
const DATE_FORMAT = { day: "2-digit", month: "2-digit", year: "numeric" } as const;
const formatDate = (when: string | number) => new Date(when).toLocaleDateString(locale(), DATE_FORMAT);

/** Empty paragraph after an inserted card, so the cursor ends up below it (see insertIntoBody). */
const AFTER_CARD = "<p style=\"margin:0;\"><br></p>";

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
  // by hand ("Verlauf"), so the draft must not be released then (see releaseDraftIfLinksGone).
  await officeAsync<void>((cb) => body.setSelectedDataAsync(data, options, cb));
  if (await bodyContains(marker)) return rememberLink(marker);

  await officeAsync<void>((cb) => body.prependAsync(data, options, cb));
  if (await bodyContains(marker)) return rememberLink(marker);

  throw new Error(t("insert.notTaken"));
}

/** Inserts the shared files as a card with download buttons (see mailCard.ts). */
async function insertLinks(results: UploadResult[]) {
  const files = results.map((r) => ({ name: r.fileName, size: r.size, url: r.shareUrl }));
  const withPassword = !isWithoutPassword(results[0]);
  // The card in the language chosen for the recipients (dates and sizes too).
  const [html, text] = inLanguage(recipientLanguage(), () => {
    const expire = formatDate(results[0].expireDate);
    const proof = cardProof(results.find((r) => r.proofId)?.proofId, true);
    return [filesCard(files, expire, formatSize, withPassword, proof), filesText(files, expire, formatSize, withPassword, proof)];
  });
  await insertIntoBody(html, text, shareToken(results[0].shareUrl));
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
 * The last step with a password in this mail (DraftState.lastStep): "Gleiches Kennwort wie zuvor" puts the next link
 * into it. Set when a link with a password was inserted, cleared when the links are deleted from the mail.
 */
let lastStep: { id: string; password: string } | null = null;

/**
 * Random: generated, visible, read-only. Own: hidden input with "show" toggle and live rule check.
 * Same: the password of the last step, visible, read-only – one password and one password mail for both.
 * None: no password field, no "Kennwort abrufbar" – a plain link without password mail.
 */
function setPasswordMode(mode: "same" | "random" | "own" | "none") {
  if (mode === "same" && !lastStep) mode = "random";
  const own = mode === "own";
  const none = mode === "none";
  const same = mode === "same";
  modeSameRow.hidden = !lastStep;
  modeSame.checked = same;
  modeRandom.checked = mode === "random";
  modeOwn.checked = own;
  modeNone.checked = none;
  passwordGroup.hidden = none;
  passwordToField.hidden = none;
  applyLinkMax();
  passwordInput.readOnly = !own;
  generateButton.hidden = own || same;
  revealButton.hidden = !own;
  rulesList.hidden = !own;
  if (none) {
    passwordInput.value = "";
  } else if (same) {
    passwordInput.type = "text";
    passwordInput.value = lastStep!.password;
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
  renderSubmit(); // an own password must meet the rules first
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
  else if (modeSame.checked) key = "password.hintSame";
  else if (perStep) key = "password.hintSeparate";
  else key = "password.hint";
  passwordHint.textContent = t(key);
  passwordHint.classList.toggle("warning", blockedByNoPassword());
}

modeSame.addEventListener("change", () => setPasswordMode("same"));
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

/**
 * After a step: the next link joins it by default ("Gleiches Kennwort wie zuvor" – one password and one password mail
 * for file, message and upload link of the same mail); "Zufällig" or "Eigenes" start a new step with a new password.
 */
function nextStepPassword() {
  if (!perStep) return;
  if (lastStep) {
    setPasswordMode("same");
  } else if (modeRandom.checked) {
    passwordInput.value = generatePassword();
  } else if (modeOwn.checked) {
    passwordInput.value = "";
    renderRules();
  }
}

/** The step for one click: the last one with "Gleiches Kennwort wie zuvor", else a new one with the shown password. */
let pendingStep: { id: string; password: string } | null = null;
function stepForClick(): string {
  pendingStep = modeSame.checked && lastStep ? { ...lastStep } : { id: newStep(), password: passwordInput.value };
  return pendingStep.id;
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
  lastStep = null;
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
  const { value, capped } = boundedNumber(expireInput, Math.min(serverConfig.defaultExpireDays, linkMaxDays()), linkMaxDays());
  expireInput.value = String(value);
  const state = await readDraftState();
  setStatus(
    sentences(
      capped ? t("expire.capped", { max: linkMaxDays() }) : "",
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
  void renderEncrypted(); // the note is only about files
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
  // As in Outlook: the button stays greyed out until everything it needs is there (the hints say what is missing).
  submitButton.disabled = inserting || blockedByNoPassword() || !insertReady();
}

/**
 * Whether the chosen kind can be inserted now – the checks of settingsProblem and the insert functions, without
 * messages or focus changes: files selected, a message within its length, a password where one is needed (an own one
 * meeting the rules) and the address or number of the chosen password channel.
 */
function insertReady(): boolean {
  if (!serverConfig) return false;
  const kind = selectedKind();
  if (kind === "files" && selectedFileCount() === 0) return false;
  // An encrypted file among them: only after "Trotzdem als Link senden" below the list (never with LABEL_ENCRYPTED=block).
  if (kind === "files" && !encryptedField.hidden && (encryptedConfirmRow.hidden || !encryptedConfirm.checked)) return false;
  if (kind === "message") {
    const text = messageInput.value.trim();
    if (!text || text.length > messageMax()) return false;
  }
  if (kind !== "files" && withoutPassword()) return false;
  if (!passwordInput.readOnly && !checkPassword(passwordInput.value).ok) return false;
  if (!withoutPassword() && channelTarget() === null) return false;
  return true;
}

// Everything insertReady looks at: re-check while the user types or chooses.
messageInput.addEventListener("input", () => renderSubmit());
passwordInput.addEventListener("input", () => renderSubmit());
passwordToInput.addEventListener("input", () => renderSubmit());
passwordSmsInput.addEventListener("input", () => renderSubmit());

/** "Gültigkeit": only the fields of the chosen kind, collapsed behind a one-line summary of the times (language and
 *  "Zustellnachweis" are in sight under "Empfänger"). */
function renderValidity() {
  if (!serverConfig) return;
  const kind = selectedKind();
  // "Kennwort abrufbar" / "Kennwort-Link gültig für" only for a password link: not without password, not by SMS (the
  // password is in the SMS itself and works as long as the link does).
  const passwordLink = !withoutPassword() && selectedChannel() !== "sms";
  // "Zustellnachweis" only for files (offered with DELIVERY_PROOF) – not for a message or an upload link.
  proofField.hidden = !serverConfig.deliveryProof || kind !== "files";
  expireField.hidden = kind === "message";
  expireHint.textContent = t(kind === "request" ? "expire.hintRequest" : "expire.hint");
  messageDaysField.hidden = kind !== "message";
  maxAccessField.hidden = !passwordLink;
  passwordDaysField.hidden = !passwordLink;
  const lifetime = kind === "message" ? sendDays(messageDaysInput).value : linkDays();
  const parts = [t(kind === "message" ? "validity.message" : "validity.link", { count: lifetime })];
  if (passwordLink) {
    // The recipients' access to the password ends at the latest with the link (the server caps it the same way).
    const days = Math.min(sendDays(passwordDaysInput).value, lifetime);
    parts.push(t("validity.password", { count: days, times: timesText(maxAccessValue()) }));
  }
  validitySummary.textContent = parts.join(" · ");
}

/**
 * Longest "Link gültig für": a link with password lasts only as long as its password in Vaultwarden (at most 30 days,
 * `passwordLinkMaxDays`); links "Ohne Kennwort" up to SHARE_MAX_EXPIRE_DAYS.
 */
const linkMaxDays = () =>
  withoutPassword() ? serverConfig.maxExpireDays : Math.min(serverConfig.maxExpireDays, serverConfig.passwordLinkMaxDays ?? serverConfig.maxExpireDays);

/** Days of the next link within the allowed range. */
const linkDays = () => boundedNumber(expireInput, Math.min(serverConfig.defaultExpireDays, linkMaxDays()), linkMaxDays()).value;

/** Keeps "Link gültig für" within the maximum of the chosen password mode (e.g. back from "Ohne Kennwort"). */
function applyLinkMax() {
  if (!serverConfig) return;
  expireInput.max = String(linkMaxDays());
  if (Number(expireInput.value) > linkMaxDays()) expireInput.value = String(linkMaxDays());
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

/** The labels of the files about to be shared (read now if the list did not show them yet). */
async function labelsOfSelection(attachments: HTMLInputElement[], files: File[]) {
  if (labelPolicy() === "off") return [];
  return Promise.all([
    ...attachments.map(async (box) => {
      const name = box.dataset.name ?? t("files.attachment");
      return { name, label: await labelOf(box.value, name, () => readAttachment(box.value)) };
    }),
    ...files.map(async (file) => ({ name: file.name, label: await labelOf(file, file.name, async () => file) })),
  ]);
}

/**
 * A protected (encrypted) file among the ticked ones: recipients outside the organisation can usually not open it,
 * whichever way it travels. A note right below the file list (#encrypted-field); LABEL_ENCRYPTED=warn needs
 * "Trotzdem als Link senden" ticked before inserting, block refuses, allow says nothing. `labels`: the ones just read
 * for inserting – otherwise those the list knows already (large attachments are only read when inserted).
 * Returns whether inserting may go on.
 */
async function renderEncrypted(labels?: Array<{ name: string; label: FileLabel | null }>): Promise<boolean> {
  let encrypted: string[] = [];
  if (labelPolicy() !== "off" && encryptedPolicy() !== "allow" && selectedKind() === "files") {
    const known = labels ?? (await knownSelectionLabels());
    encrypted = known.filter((l) => l.label?.encrypted).map((l) => l.name);
  }
  const key = encrypted.join("|");
  if (key !== encryptedShown) {
    // Other files: the confirmation was for those – ask again.
    encryptedConfirm.checked = false;
    encryptedShown = key;
  }
  encryptedField.hidden = encrypted.length === 0;
  const names = encrypted.map((name) => `„${name}“`).join(", ");
  const block = encryptedPolicy() === "block";
  encryptedText.textContent = encrypted.length
    ? t(block ? "labels.blockedNote" : "labels.encryptedNote", { names, count: encrypted.length })
    : "";
  encryptedConfirmRow.hidden = block;
  renderSubmit();
  return encrypted.length === 0 || (!block && encryptedConfirm.checked);
}

/** The labels of the ticked files that are known already (read for the list) – without reading anything new. */
async function knownSelectionLabels() {
  const boxes = [
    ...Array.from(attachmentList.querySelectorAll<HTMLInputElement>("input:checked")).map((box) => ({ key: box.value as string | File, name: box.dataset.name ?? "" })),
    ...Array.from(pickedList.querySelectorAll<HTMLInputElement>("input:checked")).map((box) => {
      const file = pickedFiles[Number(box.value)];
      return { key: file as string | File, name: file?.name ?? "" };
    }),
  ];
  return Promise.all(boxes.filter((b) => labelCache.has(b.key)).map(async (b) => ({ name: b.name, label: await labelCache.get(b.key)! })));
}

encryptedConfirm.addEventListener("change", () => renderSubmit());

/**
 * After inserting: the mail gets the highest label of its files if its own is lower (never lowered) – LABEL_POLICY
 * "inherit" sets it and says so, "suggest" offers "Übernehmen". Only labels of the user's own catalog (other
 * organisations' labels and a missing catalog: nothing to set). Returns the sentence for the message, if any.
 */
async function labelForMail(labels: Array<{ name: string; label: FileLabel | null }>): Promise<{ text: string; apply?: () => Promise<void> } | null> {
  const known = await catalog();
  if (!known || labelPolicy() === "off") return null;
  let best: { file: string; details: LabelDetails } | undefined;
  for (const { name, label } of labels) {
    const details = label?.id ? known.find((l) => l.id === label.id) : undefined;
    if (details && (!best || details.rank > best.details.rank)) best = { file: name, details };
  }
  if (!best) return null;
  const current = await mailLabel();
  if (current === null) return null;
  const currentRank = known.find((l) => l.id === current)?.rank ?? -1;
  if (best.details.rank <= currentRank) return null;
  const chosen = best;
  const apply = () => setMailLabel(chosen.details.id);
  if (labelPolicy() === "suggest") {
    return { text: t("labels.suggest", { name: chosen.file, label: chosen.details.name }), apply };
  }
  try {
    await apply();
    return { text: t("labels.applied", { name: chosen.file, label: chosen.details.name }) };
  } catch (err) {
    return { text: t("labels.applyFailed", { error: errorMessage(err) }) };
  }
}

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
  // Sensitivity labels of the files: a protected one needs a confirmation first.
  if (labelPolicy() !== "off") setStatus(t("labels.checking"), "progress");
  const labels = await labelsOfSelection(selectedAttachments, selectedFiles);
  if (!(await renderEncrypted(labels))) {
    encryptedConfirm.focus();
    return setStatus(t(encryptedPolicy() === "block" ? "labels.blockedFirst" : "labels.confirmFirst"), "error");
  }

  await runInsert(async () => {
    const results: UploadResult[] = [];
    // One step for this click: all selected files share one password and one password mail.
    const step = stepForClick();
    let failed = false;
    try {
      for (const box of selectedAttachments) {
        const name = box.dataset.name ?? t("files.attachment");
        setProgress(0);
        setStatus(t("files.preparing", { name }), "progress");
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
            // The mail gets the sensitivity of its files (or the offer to take it over).
            const forMail = await labelForMail(labels).catch(() => null);
            const apply = forMail?.apply;
            setStatus(
              sentences(
                t("submit.inserted", { count: results.length }),
                isWithoutPassword(results[0])
                  ? t("submit.noPassword", { count: results.length })
                  : t("submit.passwordFollows"),
                ...(forMail ? [forMail.text] : []),
              ),
              "success",
              apply
                ? {
                    label: t("labels.apply"),
                    run: () =>
                      void apply().then(
                        () => setStatus(t("labels.appliedShort"), "success"),
                        (err) => setStatus(t("labels.applyFailed", { error: errorMessage(err) }), "error"),
                      ),
                  }
                : undefined,
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
      setStatus(t("request.creating"), "progress");
      const result = await api<{ shareUrl: string; expireDate: string; jobId: string; proofId?: string | null }>(
        "/api/requests",
        { method: "POST", body: JSON.stringify(await shareSettingsBody(stepForClick())) },
        await getToken(),
      );
      // Same password and password mail as the files of this mail.
      await rememberShareChoice(result.jobId, undefined, result.proofId);
      const [html, text] = inLanguage(recipientLanguage(), () => {
        const expire = formatDate(result.expireDate);
        const proof = cardProof(result.proofId, true);
        return [requestCard(result.shareUrl, expire, proof), requestText(result.shareUrl, expire, proof)];
      });
      await insertIntoBody(html, text, shareToken(result.shareUrl));
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
      setStatus(t("message.creating"), "progress");
      const result = await api<{ shareUrl: string; expireDate: string; jobId: string; proofId?: string | null }>(
        "/api/messages",
        {
          method: "POST",
          body: JSON.stringify({ ...(await shareSettingsBody(stepForClick())), text, messageDays: sendDays(messageDaysInput).value }),
        },
        await getToken(),
      );
      // Same password and password mail as the files of this mail (one step).
      await rememberShareChoice(result.jobId, undefined, result.proofId);
      const [cardHtml, cardText] = inLanguage(recipientLanguage(), () => {
        const expire = formatDate(result.expireDate);
        const proof = cardProof(result.proofId, false);
        return [messageCard(result.shareUrl, expire, proof), messageText(result.shareUrl, expire, proof)];
      });
      await insertIntoBody(cardHtml, cardText, shareToken(result.shareUrl));
      messageInput.value = "";
      renderMessageCount();
      setStatus(t("message.inserted"), "success");
    } catch (err) {
      setStatus(errorMessage(err), "error");
    }
  });
}

// --- Tabs ------------------------------------------------------------------------------

/** "send" | "shares" while writing ("Senden" | "Verlauf"), "mail" | "shares" while reading. */
type Tab = "send" | "mail" | "shares";
/** The tab shown – also while the guide covers it, so closing the guide returns there. */
let currentTab: Tab = "send";

/** Shows a tab and closes the guide; `load: false` when the list is loaded elsewhere. */
function selectTab(tab: Tab, load = true) {
  currentTab = tab;
  tabSend.setAttribute("aria-selected", String(tab === "send"));
  tabMail.setAttribute("aria-selected", String(tab === "mail"));
  tabShares.setAttribute("aria-selected", String(tab === "shares"));
  showHelp(false);
  if (tab !== "send" && load) loadShares().catch((err) => setStatus(errorMessage(err), "error"));
}

/** What covers the current tab: the user guide or the settings (both in the account menu), or nothing. */
type Overlay = "help" | "settings" | "proofs" | null;

/** Shows the guide or the settings in place of the current tab; closed, that tab shows again. */
function showOverlay(overlay: Overlay) {
  helpToggle.setAttribute("aria-expanded", String(overlay === "help"));
  panelHelp.hidden = overlay !== "help";
  panelSettings.hidden = overlay !== "settings";
  panelProofs.hidden = overlay !== "proofs";
  panelSend.hidden = overlay !== null || currentTab !== "send";
  panelMail.hidden = overlay !== null || currentTab !== "mail";
  panelShares.hidden = overlay !== null || currentTab !== "shares";
  // "Aktualisieren" only where there is a link list to reload.
  refreshSharesButton.hidden = overlay !== null || currentTab === "send";
}

const showHelp = (open: boolean) => showOverlay(open ? "help" : null);

tabSend.addEventListener("click", () => selectTab("send"));
tabMail.addEventListener("click", () => selectTab("mail"));
tabShares.addEventListener("click", () => selectTab("shares"));
// "Hilfe" in the account menu opens the guide in place of the current tab; "×" closes it, the focus goes back to the
// account.
helpToggle.addEventListener("click", () => {
  showAccountMenu(false);
  showHelp(true);
  helpClose.focus(); // screen readers start in the guide
});
helpClose.addEventListener("click", () => {
  showHelp(false);
  accountToggle.focus();
});

const settingsView = initSettings({
  token: getToken,
  setStatus,
  errorMessage,
  close: () => closeSettings(),
  onChanged: () => {
    if (readMode) return;
    // New ways for "Kennwort senden an": a mail without links yet switches to the preselected way if its own is no
    // longer offered; a mail with links keeps its way (the password job already uses it).
    void readDraftState().then((state) => {
      const { channels, defaultChannel } = userSettings();
      const current = selectedChannel();
      const next = state.jobId || channels.includes(current) ? current : defaultChannel;
      applyChannelChoices(next);
      if (next !== current) showChannel(next);
      // A new default language for the recipients applies unless this mail has its own choice.
      if (!state.recipientLanguage) renderRecipientLanguages(userSettings().recipientLanguage);
    });
  },
});
// "Einstellungen" in the account menu (bottom left) opens them in place of the current tab; "×" – or saving – closes
// them, the focus goes back to the account.
settingsOpen.addEventListener("click", () => {
  showAccountMenu(false);
  showOverlay("settings");
  void settingsView.open();
  settingsClose.focus();
});
function closeSettings() {
  showOverlay(null);
  accountToggle.focus();
}
settingsClose.addEventListener("click", closeSettings);

// --- "Zustellnachweise" (account menu) ---------------------------------------------------------------------------
// Every delivery proof of the user, newest first – also those whose links expired long ago, which "Verlauf" no longer
// lists. Searched here (tracking ID, subject, recipients, file names); the seal opens the signed PDF.

interface ProofItem {
  id: string;
  createdAt: number;
  subject: string | null;
  sentAt: number | null;
  recipients: string[];
  files: string[];
  downloads: number;
  confirmed: number;
  done: boolean;
}

let proofItems: ProofItem[] = [];

function proofRow(p: ProofItem): HTMLElement {
  const row = document.createElement("div");
  row.className = "share proof-entry";
  const seal = document.createElement("button");
  seal.type = "button";
  seal.className = "btn btn-subtle btn-icon proof-seal";
  seal.innerHTML = ICONS.proof;
  seal.title = t("shares.proofLabel", { id: p.id });
  seal.setAttribute("aria-label", seal.title);
  seal.addEventListener("click", () => void openProof(p.id, seal));
  const body = document.createElement("div");
  body.className = "share-body";
  const line = (text: string, className = "share-meta") => {
    const el = document.createElement("div");
    el.className = className;
    el.textContent = text;
    return el;
  };
  body.append(
    line(p.subject || t("shares.noSubject"), "share-title"),
    line(p.sentAt
      ? t("proofs.sentTo", { date: formatDate(p.sentAt), to: p.recipients.join(", ") || "–" })
      : t("proofs.notSent", { date: formatDate(p.createdAt) })),
    line([p.files.join(", "), p.downloads > 0 ? t("proofs.downloaded", { count: p.downloads }) : t("proofs.notDownloaded")]
      .filter(Boolean).join(" · ")),
    line(`${t("proofs.case", { id: p.id })} · ${t(p.done ? "proofs.done" : "proofs.running")}`),
  );
  row.append(seal, body);
  return row;
}

/** The list for the search: tracking ID, subject, recipients and file names, case-insensitive. */
function renderProofs() {
  const query = proofsSearch.value.trim().toLowerCase();
  const hits = proofItems.filter((p) => !query || [p.id, p.subject ?? "", ...p.recipients, ...p.files].some((s) => s.toLowerCase().includes(query)));
  proofsList.replaceChildren(...hits.map(proofRow));
  proofsEmpty.hidden = hits.length > 0;
  proofsEmpty.textContent = proofItems.length ? t("proofs.noMatch") : t("proofs.none");
}

async function openProofs() {
  showAccountMenu(false);
  showOverlay("proofs");
  proofsSearch.focus();
  proofsList.replaceChildren();
  proofsEmpty.hidden = false;
  proofsEmpty.textContent = t("proofs.loading");
  try {
    proofItems = await api<ProofItem[]>("/api/proofs", {}, await getToken());
    renderProofs();
  } catch (err) {
    proofsEmpty.textContent = t("proofs.failed", { error: errorMessage(err) });
  }
}

proofsOpen.addEventListener("click", () => void openProofs());
proofsSearch.addEventListener("input", renderProofs);
proofsClose.addEventListener("click", () => {
  showOverlay(null);
  accountToggle.focus();
});
refreshSharesButton.addEventListener("click", () => {
  refreshSharesButton.disabled = true;
  refreshSharesButton.setAttribute("aria-busy", "true");
  loadShares()
    .catch((err) => setStatus(errorMessage(err), "error"))
    .then(() => {
      refreshSharesButton.disabled = false;
      refreshSharesButton.removeAttribute("aria-busy");
    });
});

// --- History ("Verlauf") -----------------------------------------------------

interface ShareItem {
  id: string;
  /** "message": a confidential message (Vaultwarden Send) – no link to copy or insert, no download count. */
  kind: "file" | "request" | "message";
  name: string;
  /** The link; null for a message (CloudShare does not keep its key). */
  url: string | null;
  /** Message: the access ID in its link (…/#/send/<accessId>/…), to find it in a mail being read. */
  accessId?: string;
  /** "Zustellnachweis": tracking ID of the mail's case – "Nachweis" opens its PDF. */
  proofId?: string | null;
  /** Subject of the mail it went out with (title of the entry); null while unknown (not sent yet, older entries). */
  subject?: string | null;
  /** Graph ID of that mail in the user's own mailbox – a click on the subject opens it. */
  mailId?: string | null;
  /** When the mail went out (ms) and to whom (To, Cc) – the tooltip of the subject; null until it is sent. */
  sentAt?: number | null;
  sentTo?: { to: string[]; cc: string[] } | null;
  sharedAt: number;
  expiration: string | null;
  downloads: { count: number; last: string | null } | null;
  received: number | null;
  /** Password-protected link (Nextcloud cannot tell random from own passwords). */
  protected?: boolean;
  /** How the password went out after sending (mail or SMS with delivery status); null before / for older links. */
  password?: {
    channel: "mail" | "sms";
    to: string;
    status: "submitted" | "delivered" | "failed" | "unknown" | null;
    at: number;
    deliveredAt: number | null;
    /** Failed SMS: seven.io did not take it at all (false: it reported "not delivered" later). */
    refused?: boolean;
  } | null;
  /** The password can still be shown and sent again ("Kennwort"). */
  passwordAvailable?: boolean;
  /** The mail with this link has not been seen sent yet (e.g. a discarded draft) – the password did not go out. */
  unsent?: boolean;
  /** When CloudShare revokes the link if the mail is still not sent (ms); null: not automatically. */
  revokeAt?: number | null;
}

/** "Kennwort per SMS an +49 170 •••• 78 – zugestellt am 16.11.2026, 14:05" for the "Verlauf" list. */
/** The way in a few words for the line ("Kennwort per Mail", "Kennwort per SMS – zugestellt"); the rest in the tooltip. */
function passwordShort(p: NonNullable<ShareItem["password"]>): string {
  if (p.channel === "mail") return t("shares.passwordMailShort");
  const status =
    p.status === "delivered" ? t("shares.smsDelivered")
    : p.status === "failed" ? t(p.refused ? "shares.smsFailed" : "shares.smsUndelivered")
    : t("shares.smsSubmitted");
  return t("shares.passwordSmsShort", { status });
}

function passwordText(p: NonNullable<ShareItem["password"]>): string {
  if (p.channel === "mail") return t("shares.passwordMail", { to: p.to });
  const status =
    p.status === "delivered"
      ? p.deliveredAt ? t("shares.smsDeliveredAt", { time: formatDateTime(p.deliveredAt) }) : t("shares.smsDelivered")
      : p.status === "failed" ? t(p.refused ? "shares.smsFailed" : "shares.smsUndelivered")
      : t("shares.smsSubmitted");
  return t("shares.passwordSms", { to: p.to, status });
}

const REQUEST_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M10 13V4m0 0L6.5 7.5M10 4l3.5 3.5M4 12v3a1 1 0 0 0 1 1h10a1 1 0 0 0 1-1v-3" fill="none" ' +
  'stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/></svg>';

/** Envelope with a lock: a confidential message. */
const MESSAGE_ICON =
  '<svg class="file-icon" viewBox="0 0 20 20" aria-hidden="true">' +
  '<path d="M16.5 9V5.5a1 1 0 0 0-1-1h-12a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1H11" fill="none" stroke="currentColor" stroke-width="1.2"/>' +
  '<path d="m2.8 5.2 6.5 5 6.9-5" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round"/>' +
  '<rect x="12.5" y="13" width="5" height="4" rx=".6" fill="none" stroke="currentColor" stroke-width="1.2"/>' +
  '<path d="M13.6 13v-1.2a1.4 1.4 0 0 1 2.8 0V13" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>';

/**
 * Opens the sent mail of an entry ("Verlauf": click on its subject). The server translates the stored mail ID into the
 * EWS ID Outlook opens (GET /api/shares/:id/mail – also after the mail was moved); displayMessageFormAsync reports a
 * failure (Mailbox 1.9), older clients get displayMessageForm.
 */
/**
 * "Nachweis": the PDF of the mail's delivery proof in the default browser – the server hands out a signed address that
 * works for a few minutes (a download from the taskpane does not work in every Outlook).
 */
async function openProof(id: string, button: HTMLButtonElement) {
  button.disabled = true;
  try {
    const { url } = await api<{ url: string }>(`/api/proofs/${encodeURIComponent(id)}/link`, { method: "POST" }, await getToken());
    const ui = Office.context.ui as unknown as { openBrowserWindow?: (url: string) => void };
    if (ui.openBrowserWindow && Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
      ui.openBrowserWindow(url);
    } else {
      window.open(url, "_blank", "noopener");
    }
  } catch (err) {
    setStatus(t("shares.proofFailed", { error: errorMessage(err) }), "error");
  } finally {
    button.disabled = false;
  }
}

async function openSentMail(share: ShareItem) {
  try {
    const { ewsId } = await api<{ ewsId: string }>(`/api/shares/${encodeURIComponent(share.id)}/mail`, {}, await getToken());
    const mailbox = Office.context.mailbox;
    if (typeof mailbox.displayMessageFormAsync === "function") {
      await officeAsync<void>((cb) => mailbox.displayMessageFormAsync(ewsId, cb));
    } else {
      mailbox.displayMessageForm(ewsId);
    }
  } catch (err) {
    setStatus(err instanceof ApiError ? errorMessage(err) : t("shares.openMailFailed"), "error");
  }
}

/** Day with year and time: "16.11.2026, 00:42 Uhr" (de) / "16/11/2026, 00:42" (en-GB) – the wording from "shares.dateTime". */
function formatDateTime(when: string | number): string {
  const d = new Date(when);
  return t("shares.dateTime", {
    date: d.toLocaleDateString(locale(), DATE_FORMAT),
    time: d.toLocaleTimeString(locale(), { hour: "2-digit", minute: "2-digit" }),
  });
}

/** Status line: downloads for files, received uploads for upload links. */
function shareState(share: ShareItem): { text: string; done: boolean } {
  if (share.kind === "request") {
    if (share.received === null) return { text: t("shares.statusUnknown"), done: false };
    return share.received > 0
      ? { text: t("shares.received", { count: share.received }), done: true }
      : { text: t("shares.noneReceived"), done: false };
  }
  if (!share.downloads) return { text: t("shares.downloadUnknown"), done: false };
  return share.downloads.count > 0
    ? {
        text: share.downloads.last
          ? t("shares.downloadedLast", { count: share.downloads.count, time: formatDateTime(share.downloads.last) })
          : t("shares.downloaded", { count: share.downloads.count }),
        done: true,
      }
    : { text: t("shares.notDownloaded"), done: false };
}

/**
 * A file name split for shortening in the middle: the start gets the "…" when the line is too short, the end – the
 * extension and a few characters before it – always stays visible. Short names are not split.
 */
function nameParts(name: string): [string, string] {
  const ext = /\.[A-Za-z0-9]{1,8}$/.exec(name)?.[0] ?? "";
  const keep = Math.min(ext.length + 4, 12);
  return name.length <= keep + 6 ? [name, ""] : [name.slice(0, name.length - keep), name.slice(-keep)];
}

/** Still valid today (Nextcloud removes expired links only once a day; the server already leaves them out). */
const isValid = (share: ShareItem) => share.expiration === null || share.expiration >= new Date().toISOString().slice(0, 10);

/** Start of a link's period, without the year if it ends in the same year: "15.11." (de) / "15/11" (en-GB). */
function shortDate(sharedAt: number, expiration: string): string {
  const from = new Date(sharedAt);
  void expiration; // always with the year now (dd.mm.yyyy)
  return formatDate(from.getTime());
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

const COPY_ICON =
  '<svg class="icon" viewBox="0 0 20 20" aria-hidden="true"><rect x="7" y="7" width="9" height="10" rx="1.5" fill="none" ' +
  'stroke="currentColor" stroke-width="1.4"/><path d="M13 4.5V4a1 1 0 0 0-1-1H5a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h.5" ' +
  'fill="none" stroke="currentColor" stroke-width="1.4"/></svg>';

/** Icons of the actions in "Verlauf" (Fluent style, 20 × 20) – also shown next to their explanation in the help. */
const icon = (paths: string) =>
  `<svg class="icon" viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.4" ` +
  `stroke-linecap="round" stroke-linejoin="round">${paths}</svg>`;
const ICONS = {
  copy: COPY_ICON,
  copied: icon('<path d="M4.5 10.5l3.5 3.5 7.5-8"/>'),
  password: icon('<circle cx="7" cy="13" r="3.5"/><path d="M9.5 10.5L16 4M13.5 6.5l2 2M11.8 8.2l1.4 1.4"/>'),
  insert: icon(
    '<path d="M8.5 11.5a3 3 0 0 0 4.2 0l2.6-2.6a3 3 0 0 0-4.2-4.2l-1 1"/>' +
      '<path d="M11.5 8.5a3 3 0 0 0-4.2 0l-2.6 2.6a3 3 0 0 0 4.2 4.2l1-1"/>',
  ),
  revoke: icon('<circle cx="10" cy="10" r="6.5"/><path d="M5.4 5.4l9.2 9.2"/>'),
  /** "Nachweis": a document with a seal (the PDF of the delivery proof). */
  proof: icon('<path d="M11.5 3H6a1 1 0 0 0-1 1v12a1 1 0 0 0 1 1h3"/><path d="M11.5 3L15 6.5V9"/><path d="M8 7h3M8 10h2"/><circle cx="14" cy="13" r="2.5"/><path d="M12.8 15.2L12.5 18l1.5-.8 1.5.8-.3-2.8"/>'),
  /** Same drawing as the refresh button in the tab bar (taskpane.html). */
  refresh: icon('<path d="M15.5 10a5.5 5.5 0 1 1-1.6-3.9M14.5 3v3.5H11"/>'),
} as const;

/** The help explains the actions with their icons: `<span data-icon="copy">` gets the icon. */
document.querySelectorAll<HTMLElement>("[data-icon]").forEach((slot) => {
  slot.innerHTML = ICONS[slot.dataset.icon as keyof typeof ICONS] ?? "";
});

interface SentPassword {
  password: string;
  /** Distribution of the sent mail – "Empfänger der Nachricht" sends the password again exactly like this. */
  recipients: { to: string[]; cc: string[]; bcc: string[] };
  /** Other address(es) the password went to ("Andere E-Mail-Adresse") – suggested again. */
  otherTo: string[];
  /** Number of the SMS (E.164, "" if none) – suggested for sending it again by SMS. */
  sms: string;
}

type ResendChannel = "recipients" | "mail" | "sms";

/** "An: a, b · Cc: c · Bcc: d" – who gets the password again with "Empfänger der Nachricht". */
function distributionText(r: SentPassword["recipients"]): string {
  const fields: Array<[string, string[]]> = [
    [t("resend.fieldTo"), r.to],
    [t("resend.fieldCc"), r.cc],
    [t("resend.fieldBcc"), r.bcc],
  ];
  return fields
    .filter(([, list]) => list.length > 0)
    .map(([label, list]) => `${label}: ${list.join(", ")}`)
    .join(" · ");
}

/**
 * "Kennwort" at a sent link: the password (read from its Vaultwarden Send by the server, only for its sender) and
 * "Kennwort senden an" – the same ways as the first time: the recipients of the message (To, Cc, Bcc as they were),
 * other addresses (each gets its own mail) or SMS; a mail only ever contains a link to the password.
 * Loaded on every opening, so it always shows the current state.
 */
async function fillPasswordPanel(panel: HTMLElement, share: ShareItem) {
  const note = document.createElement("p");
  note.className = "caption";
  note.textContent = t("resend.loading");
  panel.replaceChildren(note);
  let sent: SentPassword;
  try {
    sent = await api<SentPassword>(`/api/shares/${encodeURIComponent(share.id)}/password`, {}, await getToken());
  } catch (err) {
    note.textContent = t("resend.loadFailed", { error: errorMessage(err) });
    return;
  }

  const id = panel.id;
  panel.innerHTML =
    `<label class="label" for="${id}-value"></label>` +
    `<div class="input-group"><input type="text" id="${id}-value" class="input mono" readonly spellcheck="false" />` +
    `<button type="button" class="btn btn-icon copy-password">${COPY_ICON}</button></div>` +
    `<label class="label" for="${id}-channel"></label>` +
    `<div class="select-wrap"><select id="${id}-channel" class="input select"></select></div>` +
    `<input id="${id}-to" class="input channel-input" spellcheck="false" aria-describedby="${id}-hint" />` +
    `<p class="hint" id="${id}-hint"></p>` +
    `<button type="button" class="btn send-password"></button>`;
  const [passwordLabel, channelLabel] = Array.from(panel.querySelectorAll("label"));
  passwordLabel.textContent = t("resend.password");
  channelLabel.textContent = t("resend.channel");
  panel.querySelector<HTMLInputElement>(`#${id}-value`)!.value = sent.password;

  const copy = panel.querySelector<HTMLButtonElement>(".copy-password")!;
  copy.title = t("resend.copy");
  copy.setAttribute("aria-label", t("resend.copy"));
  copy.addEventListener("click", async () => {
    await copyText(sent.password);
    setStatus(t("resend.copied"), "success");
  });

  const channel = panel.querySelector<HTMLSelectElement>(`#${id}-channel`)!;
  // Preselected: the way the password went the first time – offered even if the settings leave it out now.
  const first: ResendChannel =
    share.password?.channel === "sms" && serverConfig.smsEnabled ? "sms" : sent.otherTo.length > 0 ? "mail" : "recipients";
  const offered = userSettings().channels as string[];
  const options: Array<[ResendChannel, string]> = [["recipients", t("channel.recipients")], ["mail", t("channel.mail")]];
  if (serverConfig.smsEnabled) options.push(["sms", t("channel.sms")]);
  for (const [value, text] of options) if (offered.includes(value) || value === first) channel.add(new Option(text, value));
  channel.value = first;

  const target = panel.querySelector<HTMLInputElement>(`#${id}-to`)!;
  const hint = panel.querySelector<HTMLParagraphElement>(`#${id}-hint`)!;
  const showChannel = () => {
    const value = channel.value as ResendChannel;
    const sms = value === "sms";
    // The recipients of the message need no input – the hint names them.
    target.hidden = value === "recipients";
    target.type = sms ? "tel" : "text";
    target.autocomplete = sms ? "tel" : "email";
    target.value = sms ? sent.sms : value === "mail" ? sent.otherTo.join(", ") : "";
    target.placeholder = t(sms ? "channel.smsPlaceholder" : "resend.mailPlaceholder");
    target.setAttribute("aria-label", t(sms ? "resend.smsLabel" : "resend.mailLabel"));
    hint.textContent =
      value === "recipients"
        ? t("resend.recipientsHint", { list: distributionText(sent.recipients) })
        : sms
          ? t("resend.smsHint") + smsCountriesHint()
          : t("resend.mailHint");
    target.removeAttribute("aria-invalid");
  };
  channel.addEventListener("change", showChannel);
  showChannel();

  /** A number from a country SMS may not go to: told when it is entered and instead of sending. */
  const smsTargetProblem = (): string | null => {
    if (channel.value !== "sms") return null;
    const number = normalizePhone(target.value);
    return number ? smsNumberProblem(number, "resend") : null;
  };
  target.addEventListener("change", () => {
    const problem = smsTargetProblem();
    if (problem) {
      target.setAttribute("aria-invalid", "true");
      setStatus(problem, "error");
    } else {
      target.removeAttribute("aria-invalid");
    }
  });

  const send = panel.querySelector<HTMLButtonElement>(".send-password")!;
  send.textContent = t("resend.send");
  send.addEventListener("click", async () => {
    const problem = smsTargetProblem();
    if (problem) {
      target.setAttribute("aria-invalid", "true");
      target.focus();
      setStatus(problem, "error");
      return;
    }
    send.disabled = true;
    setStatus(t("resend.sending"), "progress");
    try {
      const result = await api<{ channel: ResendChannel; to: string }>(
        `/api/shares/${encodeURIComponent(share.id)}/password/send`,
        { method: "POST", body: JSON.stringify({ channel: channel.value, to: target.value }) },
        await getToken(),
      );
      const sentKey = { recipients: "resend.sentRecipients", mail: "resend.sentMail", sms: "resend.sentSms" }[result.channel];
      setStatus(t(sentKey, { to: result.to }), "success");
      loadShares().catch(() => undefined); // shows the new state of the password
    } catch (err) {
      setStatus(t("resend.failed", { error: errorMessage(err) }), "error");
      send.disabled = false;
    }
  });
}

function shareRow(share: ShareItem): HTMLElement {
  const row = document.createElement("div");
  row.className = "share";
  row.dataset.id = share.id; // reading a mail, a link can be in both lists
  const state = shareState(share);
  row.innerHTML =
    // The kind at a glance: a document, the upload arrow, the envelope with a lock.
    (share.kind === "request" ? REQUEST_ICON : share.kind === "message" ? MESSAGE_ICON : FILE_ICON) +
    `<div class="share-body">` +
    `<div class="share-head"><div class="share-name"></div>` +
    // Icons only: the taskpane is narrow (Outlook sets its width); the text is in title / aria-label (and the help).
    `<div class="share-actions">` +
    `<button type="button" class="btn btn-subtle btn-icon copy">${ICONS.copy}</button>` +
    `<button type="button" class="btn btn-subtle btn-icon show-password" aria-expanded="false" hidden>${ICONS.password}</button>` +
    `<button type="button" class="btn btn-subtle btn-icon insert">${ICONS.insert}</button>` +
    `<button type="button" class="btn btn-subtle btn-icon btn-danger revoke">${ICONS.revoke}</button>` +
    `</div></div>` +
    `<div class="share-meta"><span class="share-period"></span> · <span class="share-password"></span></div>` +
    `<div class="share-meta share-unsent" hidden></div>` +
    `<div class="share-state${state.done ? " done" : ""}">` +
    `<span class="dot-state" aria-hidden="true"></span><span class="state-text"></span></div>` +
    `<div class="password-panel" hidden></div></div>`;
  const label = (button: HTMLButtonElement, text: string) => {
    button.title = text;
    button.setAttribute("aria-label", text);
  };
  label(row.querySelector<HTMLButtonElement>(".copy")!, t("shares.copyLabel"));
  label(row.querySelector<HTMLButtonElement>(".insert")!, t("shares.insertLabel"));
  label(row.querySelector<HTMLButtonElement>(".revoke")!, t("shares.revokeLabel"));
  // Too long for the line: shortened in the middle ("cloudshare-icon-va…nten.png"), so the extension stays visible.
  const [start, end] = nameParts(share.name);
  const nameStart = document.createElement("span");
  nameStart.className = "name-start";
  nameStart.textContent = start;
  const nameEnd = document.createElement("span");
  nameEnd.className = "name-end";
  nameEnd.textContent = end;
  const name = row.querySelector<HTMLElement>(".share-name")!;
  name.replaceChildren(nameStart, nameEnd);
  name.title = share.name;
  // Every entry the same way: the subject of its mail as the title (the draft's until it is sent; a click opens a sent
  // mail), below it what went out – the file with its label. Entries from before subjects were kept say so.
  {
    const title = document.createElement(share.mailId ? "button" : "div");
    title.className = "share-title";
    // No subject yet in a mail not sent: it may still come ("Noch kein Betreff").
    title.textContent =
      typeof share.subject === "string"
        ? share.subject || t(share.unsent ? "shares.noSubjectYet" : "shares.noSubject")
        : t("shares.subjectUnknown");
    title.classList.toggle("unknown", typeof share.subject !== "string");
    // Hovering the subject: when the mail went out and to whom (To, Cc – never Bcc), then how to open it.
    const sent = share.sentTo
      ? [
          t("shares.sentTo", { time: formatDateTime(share.sentAt ?? share.sharedAt), to: share.sentTo.to.join(", ") || "–" }),
          share.sentTo.cc.length ? t("shares.sentCc", { cc: share.sentTo.cc.join(", ") }) : "",
        ]
      : [];
    if (share.mailId) {
      const button = title as HTMLButtonElement;
      button.type = "button";
      button.addEventListener("click", () => void openSentMail(share));
      // That it opens the mail shows as a link on hover; screen readers hear it with the subject.
      button.setAttribute("aria-label", `${title.textContent}, ${t("shares.openMail")}`);
    }
    // The tooltip only says when and to whom it went.
    if (sent.length) title.title = sent.filter(Boolean).join("\n");
    // What went out, in plain text (the symbol on the left shows the kind): the file name, or "Dateianforderung" /
    // "Vertrauliche Nachricht".
    const item = document.createElement("div");
    item.className = "share-item";
    name.replaceWith(title);
    if (share.kind === "file") {
      item.append(name);
    } else {
      item.textContent = t(`shares.kind.${share.kind}`);
    }
    // "Zustellnachweis": a small seal after it – the tracking ID in its tooltip, a click opens the PDF ("Nachweis").
    if (share.proofId) {
      const seal = document.createElement("button");
      seal.type = "button";
      seal.className = "btn btn-subtle btn-icon proof-seal";
      seal.innerHTML = ICONS.proof;
      label(seal, t("shares.proofLabel", { id: share.proofId }));
      seal.addEventListener("click", () => void openProof(share.proofId!, seal));
      item.append(seal);
    }
    row.querySelector(".share-head")!.after(item);
  }

  // "15.11. – 29.11.2026 · Kennwort per SMS an +49 170 •••• 72 – übermittelt": the word "Kennwort" only once.
  const period = row.querySelector<HTMLElement>(".share-period")!;
  const sharedOn = formatDate(share.sharedAt);
  period.textContent = share.expiration
    ? t("shares.period", { from: shortDate(share.sharedAt, share.expiration), until: formatDate(`${share.expiration}T12:00:00`) })
    : t("shares.since", { from: sharedOn });
  period.title = [t("shares.sharedOn", { date: sharedOn }), share.expiration ? t("shares.validUntil", { date: formatDate(share.expiration) }) : ""]
    .filter(Boolean)
    .join(", ");
  const passwordLine = row.querySelector<HTMLElement>(".share-password")!;
  passwordLine.textContent = share.password
    ? passwordShort(share.password)
    : t(share.protected === false ? "shares.withoutPassword" : "shares.withPassword");
  passwordLine.title = share.password ? passwordText(share.password) : "";
  passwordLine.classList.toggle("failed", share.password?.status === "failed");

  row.querySelector<HTMLElement>(".state-text")!.textContent = state.text;
  // A message: no download state (Vaultwarden does not report it), no link to copy or insert again (no key kept).
  if (share.kind === "message") {
    row.querySelector<HTMLElement>(".share-state")!.hidden = true;
    row.querySelector<HTMLElement>(".copy")!.hidden = true;
  }
  if (share.unsent) {
    const line = row.querySelector<HTMLElement>(".share-unsent")!;
    // Downloaded already: it reached someone (e.g. sent from an Outlook without CloudShare) – only the password is missing.
    line.textContent = share.downloads?.count
      ? t("shares.unsentDownloaded")
      : share.revokeAt
        ? t("shares.unsentRevoke", { date: formatDate(share.revokeAt) })
        : t("shares.unsent");
    line.hidden = false;
  }

  // "Kennwort": show the password of a sent link and send it again (until the link expires).
  if (share.passwordAvailable) {
    const toggle = row.querySelector<HTMLButtonElement>(".show-password")!;
    const panel = row.querySelector<HTMLDivElement>(".password-panel")!;
    panel.id = `password-panel-${share.id.replace(/\W/g, "")}`;
    toggle.setAttribute("aria-controls", panel.id);
    toggle.hidden = false;
    label(toggle, t("shares.passwordLabel", { name: share.name }));
    toggle.addEventListener("click", () => {
      const open = panel.hidden === true;
      toggle.setAttribute("aria-expanded", String(open));
      panel.hidden = !open;
      if (open) void fillPasswordPanel(panel, share);
    });
  }

  const copy = row.querySelector<HTMLButtonElement>(".copy")!;
  copy.addEventListener("click", async () => {
    if (!share.url) return;
    await copyText(share.url);
    copy.innerHTML = ICONS.copied;
    label(copy, t("shares.copied"));
    setTimeout(() => {
      copy.innerHTML = ICONS.copy;
      label(copy, t("shares.copyLabel"));
    }, 2000);
  });

  // A new link for this mail with the settings of the "Senden" tab – the original password cannot be recovered.
  const insert = row.querySelector<HTMLButtonElement>(".insert")!;
  insert.hidden = readMode || share.kind === "message"; // a mail being read cannot take a link
  insert.addEventListener("click", () => {
    linkAgain(share, insert).catch(() => undefined);
  });

  // Two-step confirmation instead of a modal dialog (add-ins cannot show native dialogs).
  const revoke = row.querySelector<HTMLButtonElement>(".revoke")!;
  let armedUntil = 0;
  revoke.addEventListener("click", async () => {
    if (Date.now() > armedUntil) {
      armedUntil = Date.now() + 4000;
      // The icon becomes the word "Wirklich?" until the second click (or 4 s).
      revoke.textContent = t("shares.confirm");
      label(revoke, t("shares.confirmLabel"));
      revoke.classList.add("confirm");
      setTimeout(() => {
        if (Date.now() >= armedUntil) {
          revoke.innerHTML = ICONS.revoke;
          label(revoke, t("shares.revokeLabel"));
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
      document.querySelectorAll(`.share[data-id="${CSS.escape(share.id)}"]`).forEach((other) => other.remove());
      // A revoked link counts as expired: the server moved the file to the trash or scheduled it (CLEANUP_AFTER_DAYS).
      const fate =
        share.kind === "message"
          ? t("shares.fateMessage")
          : share.kind === "request"
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
  proofId?: string | null;
}

/**
 * "Als Link einfügen" in "Verlauf": creates a new link to the same file or upload folder with this mail's
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
    setStatus(t("shares.linking", { name: share.name }), "progress");
    const result = await api<LinkAgainResult>(
      `/api/shares/${encodeURIComponent(share.id)}/link`,
      { method: "POST", body: JSON.stringify(await shareSettingsBody(stepForClick())) },
      await getToken(),
    );
    await rememberShareChoice(result.jobId, isWithoutPassword(result), result.proofId);
    const withPassword = !isWithoutPassword(result);
    // The card in the language chosen for the recipients of this mail.
    const [html, text] = inLanguage(recipientLanguage(), () => {
      const expire = formatDate(result.expireDate);
      const proof = cardProof(result.proofId, true);
      if (result.kind === "request") return [requestCard(result.shareUrl, expire, proof), requestText(result.shareUrl, expire, proof)];
      const files = [{ name: result.fileName, size: result.size, url: result.shareUrl }];
      return [filesCard(files, expire, formatSize, withPassword, proof), filesText(files, expire, formatSize, withPassword, proof)];
    });
    await insertIntoBody(html, text, shareToken(result.shareUrl));
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

/** When the list was last loaded – it reloads when the pane comes back into view, at most every 10 s. */
let sharesLoadedAt = 0;

/** All valid links (newest first); there is no refresh button – the list loads with the tab and on return. */
async function loadShares() {
  sharesLoadedAt = Date.now();
  const shares = (await api<ShareItem[]>("/api/shares", {}, await getToken())).filter(isValid);
  if (readMode) {
    // "Nachricht": the links of the mail being read; "Verlauf": all of them.
    const tokens = await mailTokens();
    const inMail = (share: ShareItem) => tokens.has(share.url ? shareToken(share.url) : (share.accessId ?? ""));
    mailShares.replaceChildren(...shares.filter(inMail).map(shareRow));
  }
  sharesList.replaceChildren(...shares.map(shareRow));
  renderEmptyLists();
}

/** Back in the pane (other window, other app): fresh download and password states, unless a panel is open. */
function reloadSharesOnReturn() {
  const panelOpen = document.querySelector(".password-panel:not([hidden])") !== null;
  if (currentTab === "send" || helpToggle.getAttribute("aria-expanded") === "true" || panelOpen) return;
  if (document.visibilityState !== "visible" || Date.now() - sharesLoadedAt < 10_000) return;
  loadShares().catch(() => undefined);
}
window.addEventListener("focus", reloadSharesOnReturn);
document.addEventListener("visibilitychange", reloadSharesOnReturn);

/** "Noch keine Freigaben" / "keine Links in dieser Nachricht" when a list is (or became) empty. */
function renderEmptyLists() {
  noShares.hidden = sharesList.childElementCount > 0;
  mailNone.hidden = mailShares.childElementCount > 0;
}

/**
 * Tokens of the links in the mail being read – Nextcloud links end in /s/<token>, confidential messages are Send links
 * …/#/send/<accessId>/<key>. Only the user's own entries are matched against them, so other links do no harm.
 */
async function mailTokens(): Promise<Set<string>> {
  const item = Office.context.mailbox.item as Office.MessageRead | null | undefined;
  if (!item) return new Set(); // pinned pane, no mail selected
  const html = await officeAsync<string>((cb) => item.body.getAsync(Office.CoercionType.Html, cb)).catch(() => "");
  return new Set([
    ...Array.from(html.matchAll(/\/s\/([A-Za-z0-9]+)/g), (match) => match[1]),
    ...Array.from(html.matchAll(/#\/send\/([A-Za-z0-9_-]+)\//g), (match) => match[1]),
  ]);
}

/**
 * Read mode: tabs "Nachricht" (the links of the mail being read, in place of "Senden") and "Verlauf"
 * (plus "Hilfe"); follows the selected mail while the pane is pinned.
 */
async function startReadView() {
  tabSend.hidden = true;
  tabMail.hidden = false;
  selectTab("mail", false); // loaded as soon as the connection is known (showConnection)
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => {
      setStatus(""); // a message about the previous mail (e.g. a revoked link) does not belong to this one
      selectTab("mail"); // another mail: its links
    });
  } catch {
    // Pinning is not available in every Outlook version – the pane then closes with the mail anyway.
  }
  try {
    await refreshConnection(); // loads the shares once connected (showConnection)
  } catch (err) {
    setStatus(t("signInFailed", { error: errorMessage(err) }), "error", retryAction());
  }
}

/** Reloads the page unless it was reloaded for the language in this window already; true if it reloads. */
function reloadOnce(): boolean {
  const key = "cloudshare-language-reload";
  try {
    if (sessionStorage.getItem(key)) return false;
    sessionStorage.setItem(key, "1");
  } catch {
    return false; // no session storage: better stay in the old language than risk a loop
  }
  location.reload();
  return true;
}

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
    setStatus(t("unreachable", { url: API_BASE }), "error", retryAction());
    return;
  }
  // Personal settings (ways to send the password, signature, language). Without them the usual ones apply.
  try {
    if (await loadSettings(await getToken())) {
      // The language was chosen on another computer: start again in it – once, so a failed save cannot loop.
      if (reloadOnce()) return;
    }
  } catch {
    // e.g. sign-in not possible yet – the connection check below reports that
  }
  // "Zustellnachweise" in the account menu – also while reading a mail – only where the server has the delivery proof.
  proofsOpen.hidden = !serverConfig.deliveryProof;
  if (readMode) return startReadView();

  const state = await readDraftState();
  expireInput.value = String(serverConfig.defaultExpireDays);
  // "Zustellnachweis": offered only when the server has it (DELIVERY_PROOF); never ticked unless this mail has it.
  proofField.hidden = !serverConfig.deliveryProof;
  // The guide explains it only where it exists.
  for (const id of ["help-proof", "help-shares-proof"]) $<HTMLElement>(id).hidden = !serverConfig.deliveryProof;
  $<HTMLElement>("help-files-labels").hidden = serverConfig.labelPolicy === "off";
  proofInput.checked = state.proof === true;
  applyLinkMax();
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
    // The pane opened again on a mail with links: the next link joins the last step by default.
    lastStep = state.lastStep ?? null;
    setPasswordMode(lastStep ? "same" : "random");
  } else {
    setPasswordMode("random");
  }
  // "Kennwort senden an": SMS only when the server has it configured; restore the choice of this draft.
  if (!serverConfig.smsEnabled) channelSmsOption.remove();
  renderMessageCount();
  passwordToInput.value = state.passwordTo ?? "";
  passwordSmsInput.value = state.passwordSms ?? "";
  // The way this mail already uses; a new mail starts with the way preselected in the settings.
  const preferred = userSettings().defaultChannel;
  const channel: Channel =
    state.passwordSms && serverConfig.smsEnabled
      ? "sms"
      : state.passwordTo
        ? "mail"
        : state.jobId || (preferred === "sms" && !serverConfig.smsEnabled)
          ? "recipients"
          : preferred;
  applyChannelChoices(channel);
  showChannel(channel);
  // "Sprache für die Empfänger": this mail's choice, else the one from the settings.
  renderRecipientLanguages(state.recipientLanguage ?? userSettings().recipientLanguage);
  renderKind();
  // Outlook has no "body changed" event: check when the pane gets focus and every few seconds while locked.
  await releaseIfLinksDeleted();
  window.addEventListener("focus", () => {
    void releaseIfLinksDeleted();
    void syncDraftSubject();
  });
  setInterval(() => {
    void releaseIfLinksDeleted();
    void syncDraftSubject();
  }, 4000);

  await renderAttachments().catch(() => undefined);
  composeItem().addHandlerAsync(Office.EventType.AttachmentsChanged, () => {
    renderAttachments().catch(() => undefined);
  });
  // Recipients decide which size applies: preselect again when it changes (internal ↔ external).
  composeItem().addHandlerAsync(Office.EventType.RecipientsChanged, () => {
    void attachmentThreshold(serverConfig).then((mb) => {
      if (mb !== shownThreshold) renderAttachments().catch(() => undefined);
    });
  });

  try {
    await refreshConnection();
  } catch (err) {
    setStatus(t("signInFailed", { error: errorMessage(err) }), "error", retryAction());
  }
});
