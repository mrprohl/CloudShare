/**
 * Personal settings ("Einstellungen" in the account menu): notices, the ways "Kennwort senden an" offers and the one it preselects, the language preselected for the recipients,
 * the language of the add-in and an own text signature for the password mails. Stored on the server per user
 * (server/src/settings.ts, /api/settings); the language is also kept in roamingSettings (SETTINGS_CACHE_KEY), so the
 * pane and the event handlers start in it without a request. What the organisation decides (notices switched off,
 * SMS without seven.io, SMS countries and blocked numbers) is not offered here. Every change is saved at once
 * (no "Speichern"); notices and signature are in sight, the ways for the password, the languages and "Zurücksetzen"
 * under "Weitere Einstellungen".
 */
import { desiredLanguage, language, languageChoices, SETTINGS_CACHE_KEY, t } from "../shared/i18n";
import { api, type Channel, type SettingsResponse, type UserSettings } from "../shared/office";

const CHANNELS: Channel[] = ["recipients", "mail", "sms"];
/** Profile fields a signature may use (server/src/graph.ts, Profile; filled by server/src/signature.ts). */
const PLACEHOLDERS = ["displayName", "givenName", "surname", "jobTitle", "department", "company", "email", "phone", "mobile", "office"];
const channelLabel = (c: Channel) => t(c === "recipients" ? "channel.recipients" : c === "mail" ? "channel.mail" : "channel.sms");

let current: SettingsResponse | null = null;

/** The settings that apply; before they are loaded (or if loading failed) the organisation's usual ones. */
export function userSettings(): UserSettings {
  return (
    current?.settings ?? {
      downloadNotice: true,
      uploadNotice: true,
      smsNotice: true,
      proofReminder: true,
      channels: CHANNELS,
      defaultChannel: "recipients",
      language: "auto",
      recipientLanguage: language(),
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
  const recipientLanguageSelect = $<HTMLSelectElement>("settings-recipient-language");
  const signatureFormat = $<HTMLSelectElement>("settings-signature-format");
  const signatureText = $<HTMLTextAreaElement>("settings-signature-text");
  const signatureHint = $<HTMLParagraphElement>("settings-signature-hint");
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

  // Placeholders of the profile behind the "i": a click inserts one where the cursor is.
  const placeholdersToggle = $<HTMLButtonElement>("settings-placeholders-toggle");
  const placeholders = $<HTMLDivElement>("settings-placeholders");
  const placeholderList = $<HTMLUListElement>("settings-placeholder-list");
  const showPlaceholders = (open: boolean) => {
    placeholders.hidden = !open;
    placeholdersToggle.setAttribute("aria-expanded", String(open));
  };
  placeholderList.replaceChildren(
    ...PLACEHOLDERS.map((name) => {
      const item = document.createElement("li");
      const insert = document.createElement("button");
      insert.type = "button";
      insert.className = "placeholder";
      insert.textContent = `{${name}}`;
      insert.setAttribute("aria-label", t("settings.placeholders.insert", { placeholder: `{${name}}`, label: t(`settings.placeholders.${name}`) }));
      insert.addEventListener("click", () => {
        const start = signatureText.selectionStart ?? signatureText.value.length;
        const end = signatureText.selectionEnd ?? start;
        signatureText.setRangeText(`{${name}}`, start, end, "end");
        signatureText.focus();
        signatureText.dispatchEvent(new Event("input")); // saved like typed text
      });
      const label = document.createElement("span");
      label.textContent = t(`settings.placeholders.${name}`);
      item.append(insert, label);
      return item;
    }),
  );
  placeholdersToggle.addEventListener("click", () => showPlaceholders(placeholders.hidden === true));

  const showSignatureFormat = () => {
    const format = signatureFormat.value;
    signatureText.hidden = format !== "text";
    placeholdersToggle.hidden = format !== "text";
    if (format !== "text") showPlaceholders(false);
    signatureHint.textContent = t(format === "text" ? "settings.signature.hint.text" : "settings.signature.hint.none");
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
    if (available.proofReminder) {
      noticeRows.push(checkbox("proofReminder", "1", t("settings.notices.proofReminder"), settings.proofReminder, t("settings.notices.proofReminderHint")));
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
    // "Sprache für die Empfänger": the languages the server offers, by their own name.
    recipientLanguageSelect.replaceChildren(
      ...languageChoices()
        .filter((l) => available.languages.includes(l.code))
        .map((l) => new Option(l.name, l.code)),
    );
    recipientLanguageSelect.value = settings.recipientLanguage;

    const signature = settings.signature;
    signatureFormat.value = signature?.format === "text" ? "text" : "none";
    signatureText.value = signature?.format === "text" ? signature.content : "";
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

  /** Takes the stored settings; a new language reloads the page. `refill`: show them in the form (after "Zurücksetzen"). */
  const apply = async (response: SettingsResponse, message: string, refill: boolean) => {
    current = response;
    await cacheLanguage(response.settings);
    if (desiredLanguage() !== language()) {
      location.reload(); // the page and every text in the new language
      return;
    }
    // While typing, the form stays as it is (a signature still empty is not switched back to the organisation's).
    if (refill) fill(response);
    options.onChanged();
    options.setStatus(message, "success");
  };

  /** What the form shows, as the body of PUT /api/settings; null while no way for the password is ticked. */
  const formSettings = (): Partial<UserSettings> | null => {
    const channels = checkedChannels();
    if (channels.length === 0) return null;
    // Plain text; an empty line starts a new paragraph in the password mail.
    const text = signatureText.value.trim();
    const signature: UserSettings["signature"] = signatureFormat.value === "text" && text ? { format: "text", content: text } : null;
    const body: Partial<UserSettings> = {
      channels,
      defaultChannel: defaultChannel.value as Channel,
      language: languageSelect.value,
      recipientLanguage: recipientLanguageSelect.value,
      signature,
    };
    for (const input of notices.querySelectorAll<HTMLInputElement>("input")) {
      (body as Record<string, unknown>)[input.name] = input.checked;
    }
    return body;
  };

  // Saved at once: every change of a box or a list, the signature after a pause in typing and when leaving it. One
  // save at a time; a change meanwhile is saved right after with what the form then shows.
  let saving: Promise<void> = Promise.resolve();
  let lastSaved = "";
  const saveNow = () => {
    saving = saving.then(async () => {
      const body = formSettings();
      if (!body) {
        options.setStatus(t("settings.channelsNone"), "error");
        return;
      }
      const json = JSON.stringify(body);
      if (json === lastSaved) return;
      try {
        const response = await api<SettingsResponse>("/api/settings", { method: "PUT", body: json }, await options.token());
        lastSaved = json;
        await apply(response, t("settings.saved"), false);
      } catch (err) {
        options.setStatus(t("settings.saveFailed", { error: options.errorMessage(err) }), "error");
      }
    });
    return saving;
  };
  let typing: ReturnType<typeof setTimeout> | undefined;
  form.addEventListener("submit", (event) => event.preventDefault());
  form.addEventListener("change", () => {
    clearTimeout(typing);
    void saveNow();
  });
  signatureText.addEventListener("input", () => {
    clearTimeout(typing);
    typing = setTimeout(() => void saveNow(), 1000);
  });

  reset.addEventListener("click", async () => {
    reset.disabled = true;
    try {
      const response = await api<SettingsResponse>("/api/settings", { method: "DELETE" }, await options.token());
      lastSaved = "";
      await apply(response, t("settings.resetDone"), true);
    } catch (err) {
      options.setStatus(t("settings.saveFailed", { error: options.errorMessage(err) }), "error");
    } finally {
      reset.disabled = false;
    }
  });

  return { open };
}
