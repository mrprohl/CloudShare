/**
 * Personal settings (⚙ in the account line): notices, the ways "Kennwort senden an" offers and the one it preselects,
 * the language of the add-in and an own signature for the link card. Stored on the server per user
 * (server/src/settings.ts, /api/settings); the language is also kept in roamingSettings (SETTINGS_CACHE_KEY), so the
 * pane and the event handlers start in it without a request. What the organisation decides (notices switched off,
 * SMS without seven.io, SMS countries and blocked numbers) is not offered here.
 */
import { desiredLanguage, language, languageChoices, SETTINGS_CACHE_KEY, t } from "../shared/i18n";
import { api, type Channel, type SettingsResponse, type UserSettings } from "../shared/office";
import { cleanInPlace, signatureFromNodes, signatureHasContent, signatureNodes } from "./signature";

const CHANNELS: Channel[] = ["recipients", "mail", "sms"];
const channelLabel = (c: Channel) => t(c === "recipients" ? "channel.recipients" : c === "mail" ? "channel.mail" : "channel.sms");

let current: SettingsResponse | null = null;

/** The settings that apply; before they are loaded (or if loading failed) the organisation's usual ones. */
export function userSettings(): UserSettings {
  return (
    current?.settings ?? {
      downloadNotice: true,
      uploadNotice: true,
      smsNotice: true,
      channels: CHANNELS,
      defaultChannel: "recipients",
      language: "auto",
      signature: null,
    }
  );
}

/**
 * Keeps the language in roamingSettings for the next start (i18n.ts reads it) – resolved once Outlook saved it (at
 * most 3 s), so a reload right after finds it. Best effort.
 */
function cacheLanguage(settings: UserSettings): Promise<void> {
  return new Promise((resolve) => {
    try {
      const roaming = Office.context.roamingSettings;
      const cached = roaming.get(SETTINGS_CACHE_KEY) as { language?: string } | undefined;
      if (cached?.language === settings.language) return resolve();
      roaming.set(SETTINGS_CACHE_KEY, { language: settings.language });
      setTimeout(resolve, 3000);
      roaming.saveAsync(() => resolve());
    } catch {
      resolve(); // no roamingSettings – the language then follows Outlook until the settings are loaded
    }
  });
}

/**
 * Loads the settings. Returns true if the language of the pane must change (chosen on another computer) – the
 * caller reloads the page then.
 */
export async function loadSettings(token: string): Promise<boolean> {
  current = await api<SettingsResponse>("/api/settings", {}, token);
  await cacheLanguage(current.settings);
  return desiredLanguage() !== language(); // language(): the page's, desiredLanguage(): the settings' / Outlook's
}

interface Options {
  token: () => Promise<string>;
  setStatus: (text: string, kind?: "success" | "error") => void;
  errorMessage: (err: unknown) => string;
  /** The settings changed (not the language – the page reloads for that). */
  onChanged: () => void;
}

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** Wires the settings form (panel-settings in taskpane.html); called once when taskpane.ts starts. */
export function initSettings(options: Options) {
  const form = $<HTMLFormElement>("settings-form");
  const loading = $<HTMLParagraphElement>("settings-loading");
  const channelBoxes = $<HTMLDivElement>("settings-channels");
  const defaultChannel = $<HTMLSelectElement>("settings-default-channel");
  const noticesGroup = $<HTMLFieldSetElement>("settings-notices-group");
  const notices = $<HTMLDivElement>("settings-notices");
  const languageSelect = $<HTMLSelectElement>("settings-language");
  const signatureFormat = $<HTMLSelectElement>("settings-signature-format");
  const signatureText = $<HTMLTextAreaElement>("settings-signature-text");
  const signatureHtml = $<HTMLDivElement>("settings-signature-html");
  const signatureHint = $<HTMLParagraphElement>("settings-signature-hint");
  const save = $<HTMLButtonElement>("settings-save");
  const reset = $<HTMLButtonElement>("settings-reset");

  const checkbox = (name: string, value: string, label: string, checked: boolean, hint?: string) => {
    const row = document.createElement("label");
    row.className = "checkbox";
    row.innerHTML = `<input type="checkbox" /><span class="box" aria-hidden="true"></span><span class="checkbox-text"></span>`;
    const input = row.querySelector("input")!;
    input.name = name;
    input.value = value;
    input.checked = checked;
    const text = row.querySelector<HTMLSpanElement>(".checkbox-text")!;
    text.textContent = label;
    if (hint) {
      const small = document.createElement("span");
      small.className = "hint checkbox-hint";
      small.textContent = hint;
      text.append(document.createElement("br"), small);
    }
    return row;
  };

  const checkedChannels = () =>
    Array.from(channelBoxes.querySelectorAll<HTMLInputElement>("input:checked")).map((input) => input.value as Channel);

  /** "Vorbelegt": only the ways that are offered. */
  const renderDefaultChannel = (wanted: Channel) => {
    const offered = checkedChannels();
    defaultChannel.replaceChildren(...offered.map((c) => new Option(channelLabel(c), c)));
    defaultChannel.value = offered.includes(wanted) ? wanted : offered[0] ?? "";
    defaultChannel.disabled = offered.length === 0;
  };

  const showSignatureFormat = () => {
    const format = signatureFormat.value;
    signatureText.hidden = format !== "text";
    signatureHtml.hidden = format !== "html";
    signatureHint.textContent = t(`settings.signature.hint.${format === "text" ? "text" : format === "html" ? "html" : "none"}`);
  };

  /** Fills the form with `response`. */
  const fill = (response: SettingsResponse) => {
    const { settings, available } = response;
    channelBoxes.replaceChildren(
      ...available.channels.map((c) => checkbox("channel", c, channelLabel(c), settings.channels.includes(c))),
    );
    renderDefaultChannel(settings.defaultChannel);

    const noticeRows: HTMLElement[] = [];
    if (available.downloadNotice) {
      noticeRows.push(checkbox("downloadNotice", "1", t("settings.notices.download"), settings.downloadNotice, t("settings.notices.downloadHint")));
    }
    if (available.uploadNotice) {
      noticeRows.push(checkbox("uploadNotice", "1", t("settings.notices.upload"), settings.uploadNotice, t("settings.notices.uploadHint")));
    }
    if (available.smsNotice) {
      noticeRows.push(checkbox("smsNotice", "1", t("settings.notices.sms"), settings.smsNotice, t("settings.notices.smsHint")));
    }
    notices.replaceChildren(...noticeRows);
    noticesGroup.hidden = noticeRows.length === 0;

    languageSelect.replaceChildren(
      new Option(t("settings.language.auto"), "auto"),
      ...languageChoices()
        .filter((l) => available.languages.includes(l.code))
        .map((l) => new Option(l.name, l.code)),
    );
    languageSelect.value = settings.language;

    const signature = settings.signature;
    signatureFormat.value = signature ? signature.format : "none";
    signatureText.value = signature?.format === "text" ? signature.content : "";
    signatureHtml.replaceChildren(...(signature?.format === "html" ? signatureFromNodes(signature.nodes, document) : []));
    showSignatureFormat();
  };

  /** Shown with the panel: the current settings, loaded again (another computer may have changed them). */
  async function open() {
    form.hidden = !current;
    loading.hidden = Boolean(current);
    if (current) fill(current);
    try {
      current = await api<SettingsResponse>("/api/settings", {}, await options.token());
      fill(current);
      form.hidden = false;
      loading.hidden = true;
    } catch (err) {
      if (!current) loading.textContent = t("settings.loadFailed", { error: options.errorMessage(err) });
    }
  }

  channelBoxes.addEventListener("change", () => renderDefaultChannel(defaultChannel.value as Channel));
  signatureFormat.addEventListener("change", showSignatureFormat);

  // Pasted from Outlook: the browser inserts the formatted text (it already drops scripts and event handlers), then
  // the allow-list of signature.ts cleans the editor in place – the clipboard is never parsed as a markup string.
  signatureHtml.addEventListener("input", (event) => {
    const type = (event as InputEvent).inputType ?? "";
    if (type.startsWith("insertFrom")) cleanInPlace(signatureHtml);
  });

  const apply = async (response: SettingsResponse, message: string) => {
    current = response;
    await cacheLanguage(response.settings);
    if (desiredLanguage() !== language()) {
      location.reload(); // the page and every text in the new language
      return;
    }
    fill(response);
    options.onChanged();
    options.setStatus(message, "success");
  };

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const channels = checkedChannels();
    if (channels.length === 0) {
      options.setStatus(t("settings.channelsNone"), "error");
      return;
    }
    const format = signatureFormat.value;
    let signature: UserSettings["signature"] = null;
    if (format === "text" && signatureText.value.trim()) {
      signature = { format: "text", content: signatureText.value.trim() };
    } else if (format === "html") {
      cleanInPlace(signatureHtml); // typed or formatted by hand after the last paste
      if (signatureHasContent(signatureHtml)) signature = { format: "html", nodes: signatureNodes(signatureHtml) };
    }
    const body: Partial<UserSettings> = {
      channels,
      defaultChannel: defaultChannel.value as Channel,
      language: languageSelect.value,
      signature,
    };
    for (const input of notices.querySelectorAll<HTMLInputElement>("input")) {
      (body as Record<string, unknown>)[input.name] = input.checked;
    }
    save.disabled = true;
    try {
      const response = await api<SettingsResponse>("/api/settings", { method: "PUT", body: JSON.stringify(body) }, await options.token());
      await apply(response, t("settings.saved"));
    } catch (err) {
      options.setStatus(t("settings.saveFailed", { error: options.errorMessage(err) }), "error");
    } finally {
      save.disabled = false;
    }
  });

  reset.addEventListener("click", async () => {
    reset.disabled = true;
    try {
      const response = await api<SettingsResponse>("/api/settings", { method: "DELETE" }, await options.token());
      await apply(response, t("settings.resetDone"));
    } catch (err) {
      options.setStatus(t("settings.saveFailed", { error: options.errorMessage(err) }), "error");
    } finally {
      reset.disabled = false;
    }
  });

  return { open };
}
