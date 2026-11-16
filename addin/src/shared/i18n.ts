/**
 * Texts of the add-in in the language of the Outlook client (Office.context.displayLanguage), English when the server
 * has no file for it. The server replaces the marker below with the "addin" texts of every language
 * (locales/<language>.json, see server/src/i18n.ts) when the container starts, so the bundles need no extra request –
 * important for the send handler in classic Outlook's JavaScript-only runtime. The static texts of taskpane.html come
 * as one page per language (taskpane.<language>.html, see taskpane/language.ts).
 */

interface Language {
  locale: string;
  /** Name of the language in itself ("Deutsch", "English") – for the language choice in the settings. */
  name?: string;
  messages: Record<string, string>;
}

/**
 * roamingSettings key of the user's settings as last loaded (taskpane, settings) – read here for the language, so the
 * page and the event handlers have it at once, before any request.
 */
export const SETTINGS_CACHE_KEY = "cloudshare-settings";

interface Bundle {
  /** Used when the client's language is not available ("en"). */
  fallback: string;
  languages: Record<string, Language>;
}

const INJECTED = "%%CLOUDSHARE_I18N%%";

const EMPTY: Bundle = { fallback: "de", languages: { de: { locale: "de-DE", messages: {} } } };

function parse(): Bundle {
  try {
    // Not replaced (unit tests): keys are shown as they are until setMessages() is called.
    if (INJECTED.charAt(0) !== "{") return EMPTY;
    const parsed = JSON.parse(INJECTED) as Bundle;
    return parsed && parsed.languages && parsed.languages[parsed.fallback] ? parsed : EMPTY;
  } catch {
    return EMPTY;
  }
}

let bundle = parse();
/** Set when the language must not follow the client (tests, a taskpane page that could not be switched). */
let forced: string | undefined;

/** Tests only: use these texts. */
export function setMessages(messages: Record<string, string>, locale = "de-DE"): void {
  const lang = locale.slice(0, 2).toLowerCase();
  bundle = { fallback: lang, languages: { [lang]: { locale, messages } } };
  forced = lang;
}

/** Tests only: use these languages and follow the client again. */
export function setLanguages(languages: Record<string, Language>, fallback = "en"): void {
  bundle = { fallback, languages };
  forced = undefined;
}

/** Languages the add-in has texts for. */
export const availableLanguages = (): string[] => Object.keys(bundle.languages).sort();

/**
 * The language for a client language tag: the exact one ("pt-br"), else the one without region ("de-AT" → "de"),
 * else the fallback (English). Same rule as the server (server/src/i18n.ts, pickLanguage).
 */
export function pickLanguage(tag: string | undefined, available: string[] = availableLanguages()): string {
  const wanted = (tag || "").trim().toLowerCase();
  if (available.indexOf(wanted) >= 0) return wanted;
  const base = wanted.split("-")[0];
  if (available.indexOf(base) >= 0) return base;
  return available.indexOf(bundle.fallback) >= 0 ? bundle.fallback : available[0];
}

/** Display language of the Outlook client ("de-DE"); the browser's before Office.js is ready. */
function clientLanguage(): string | undefined {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.displayLanguage) return Office.context.displayLanguage;
  } catch {
    // Office.context throws in some hosts before initialization
  }
  return typeof navigator !== "undefined" ? navigator.language : undefined;
}

/** Language chosen in the settings (not "auto") and available, or undefined. */
function preferredLanguage(): string | undefined {
  try {
    const cached = Office.context.roamingSettings.get(SETTINGS_CACHE_KEY) as { language?: unknown } | undefined;
    const lang = cached && typeof cached.language === "string" ? cached.language : "";
    return lang && lang !== "auto" && bundle.languages[lang] ? lang : undefined;
  } catch {
    return undefined; // no Office.js (yet) or no roamingSettings
  }
}

/** The language the add-in should show: the one chosen in the settings, else the client's. */
export const desiredLanguage = (): string => preferredLanguage() || pickLanguage(clientLanguage());

/** The language the add-in shows ("de", "en", …): fixed for a taskpane page (forceLanguage), else desiredLanguage. */
export const language = (): string => forced || desiredLanguage();

/** The languages for a choice: code and own name, sorted by name. */
export const languageChoices = (): Array<{ code: string; name: string }> =>
  availableLanguages()
    .map((code) => ({ code, name: bundle.languages[code].name || code }))
    .sort((a, b) => a.name.localeCompare(b.name));

/** Keeps the add-in in `lang` whatever the client says (a taskpane page that stays in its language). */
export function forceLanguage(lang: string): void {
  if (bundle.languages[lang]) forced = lang;
}

const current = (): Language => bundle.languages[language()] || bundle.languages[bundle.fallback];

/** BCP 47 locale for dates and numbers ("de-DE", "en-GB"). */
export const locale = () => current().locale;

/**
 * Text for `key` with {name} placeholders filled. With a numeric `count`, "key_one" (count 1) or "key_other" is
 * used if the language has it.
 */
export function t(key: string, vars: Record<string, string | number> = {}): string {
  const messages = current().messages;
  const count = vars.count;
  let text: string | undefined;
  if (typeof count === "number") text = messages[count === 1 ? `${key}_one` : `${key}_other`];
  if (text === undefined) text = messages[key];
  if (text === undefined) text = key;
  return text.replace(/\{(\w+)\}/g, (match, name: string) =>
    Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
  );
}
