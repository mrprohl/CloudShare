/**
 * Languages. Every text users or recipients see comes from a language file: locales/<language>.json in the
 * repository (shipped in the image as /app/locales) and optionally CUSTOM_LOCALES_DIR (default /config/locales,
 * mounted from env/config/locales) for an own language or single changed texts without a new image.
 *
 * The add-in follows the language of the Outlook client: every available language goes into the add-in files, the
 * add-in picks the one of Office.context.displayLanguage (English if there is none) and sends it with every request
 * (header X-CloudShare-Language), so the error messages it shows come in that language too. LANGUAGE (default "en")
 * is the language of everything CloudShare sends itself (mails, SMS) and the first guess of the taskpane page.
 * Missing texts fall back to English, then to the key itself, so a partly translated file still works. Sections:
 *   addin      taskpane, Smart Alerts, info bar, link cards – injected into the add-in files at startup
 *   manifest   texts of manifest.xml (description, ribbon) – English as default, one Override per other language
 *   server     error messages the taskpane shows – in the language of the request
 *   mail       password mail – LANGUAGE
 *   templates  built-in texts of templates.ts (SMS, password mail, download confirmation) – LANGUAGE
 *
 * Placeholders are {name}. Plural forms: "key_one" (count 1) and "key_other"; t(key, { count }) picks one.
 */
import { AsyncLocalStorage } from "node:async_hooks";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { config } from "./config.js";

export type Section = Record<string, string>;
export interface LocaleFile {
  _meta: { name: string; locale: string };
  addin: Section;
  manifest: Section;
  server: Section;
  mail: Section;
  templates: Section;
}

export const SECTIONS = ["addin", "manifest", "server", "mail", "templates"] as const;
export const FALLBACK_LANGUAGE = "en";
const LANGUAGE_PATTERN = /^[a-z]{2,3}(-[a-z0-9]{2,8})?$/;

const here = path.dirname(fileURLToPath(import.meta.url));
/** Built-in language files: /app/locales in the image, <repo>/locales when run from server/src or server/dist. */
export const builtinLocalesDir =
  process.env.LOCALES_DIR ??
  [path.resolve(here, "../locales"), path.resolve(here, "../../locales")].find((dir) => existsSync(dir)) ??
  path.resolve(here, "../locales");
export const customLocalesDir = process.env.CUSTOM_LOCALES_DIR ?? "/config/locales";

function readLocale(dir: string, language: string): Partial<LocaleFile> | undefined {
  const file = path.join(dir, `${language}.json`);
  if (!existsSync(file)) return undefined;
  try {
    return JSON.parse(readFileSync(file, "utf8")) as Partial<LocaleFile>;
  } catch (err) {
    throw new Error(`Language file ${file} is not valid JSON: ${(err as Error).message}`);
  }
}

/** Languages available in the built-in and the custom directory (file names without .json). */
export function availableLanguages(dirs = [builtinLocalesDir, customLocalesDir]): string[] {
  const names = new Set<string>();
  for (const dir of dirs) {
    try {
      for (const file of readdirSync(dir)) if (file.endsWith(".json")) names.add(file.slice(0, -5));
    } catch {
      // directory missing – fine for the custom one
    }
  }
  return [...names].sort();
}

/**
 * The complete texts of `language`: English as the base, then the built-in file of the language, then the custom
 * one – each level may contain only some texts. Throws for an unknown language (stops the start with a clear message).
 */
export function loadLocale(language: string, dirs = { builtin: builtinLocalesDir, custom: customLocalesDir }): LocaleFile {
  const lang = language.trim().toLowerCase();
  if (!LANGUAGE_PATTERN.test(lang)) throw new Error(`LANGUAGE "${language}" is not a language code like "de" or "en"`);
  const levels = [readLocale(dirs.builtin, FALLBACK_LANGUAGE), readLocale(dirs.builtin, lang), readLocale(dirs.custom, lang)];
  if (!levels[1] && !levels[2]) {
    const available = availableLanguages([dirs.builtin, dirs.custom]).join(", ") || "none";
    throw new Error(`LANGUAGE "${lang}": no file ${lang}.json in ${dirs.builtin} or ${dirs.custom} (available: ${available})`);
  }
  const result: LocaleFile = {
    _meta: { name: lang, locale: lang },
    addin: {},
    manifest: {},
    server: {},
    mail: {},
    templates: {},
  };
  for (const level of levels) {
    if (!level) continue;
    if (level._meta) result._meta = { ...result._meta, ...level._meta };
    for (const section of SECTIONS) Object.assign(result[section], level[section] ?? {});
  }
  // The language's own locale ("de-DE"), not English's when only the custom file has texts.
  if (!levels[1]?._meta && !levels[2]?._meta) result._meta = { name: lang, locale: lang };
  return result;
}

export const language = config.language;
let active: LocaleFile = loadLocale(language);

/** All languages the add-in can show, with LANGUAGE and English always among them. */
function loadAll(): Record<string, LocaleFile> {
  const all: Record<string, LocaleFile> = { [language]: active };
  for (const lang of availableLanguages()) {
    if (all[lang] || !LANGUAGE_PATTERN.test(lang)) continue;
    try {
      all[lang] = loadLocale(lang);
    } catch {
      // an unreadable custom file only stops the start when it is LANGUAGE itself
    }
  }
  if (!all[FALLBACK_LANGUAGE]) all[FALLBACK_LANGUAGE] = loadLocale(FALLBACK_LANGUAGE);
  return all;
}
let all: Record<string, LocaleFile> = loadAll();

/** Tests only: switch the language (and use only it, besides English). */
export function setLocale(file: LocaleFile, languages?: Record<string, LocaleFile>): void {
  active = file;
  if (languages) all = languages;
}

/** Languages the add-in can show ("de", "en", …). */
export const addinLanguages = () => Object.keys(all).sort();

/**
 * The language for a client language tag ("de-AT" → "de", "pt-BR" → "pt-br" or "pt"): the exact one, else the one
 * without region, else English – as the add-in does (addin/src/shared/i18n.ts, pickLanguage).
 */
export function pickLanguage(tag: unknown, available: readonly string[] = addinLanguages()): string {
  const wanted = typeof tag === "string" ? tag.trim().toLowerCase() : "";
  if (available.includes(wanted)) return wanted;
  const base = wanted.split("-")[0];
  if (available.includes(base)) return base;
  return available.includes(FALLBACK_LANGUAGE) ? FALLBACK_LANGUAGE : language;
}

/** Language of the request being handled: the add-in's language for the "server" texts (error messages). */
const requestLanguage = new AsyncLocalStorage<string>();

/**
 * Runs `fn` (and everything it starts) with the "server" texts in the language of `tag`. Without a tag (an add-in
 * from before this header, webhooks, scripts) the texts stay in LANGUAGE.
 */
export function withRequestLanguage(tag: unknown, fn: () => void): void {
  const given = typeof tag === "string" && tag.trim() !== "";
  requestLanguage.run(given ? pickLanguage(tag) : language, fn);
}

/**
 * Language of what the recipients get – the password mail and SMS of a step, in the language chosen when the link was
 * inserted ("Sprache für die Empfänger"). Notes to the sender stay in LANGUAGE.
 */
const recipientLanguage = new AsyncLocalStorage<string>();

/** A language the server has texts for, or undefined ("" / unknown / not a string). */
export const knownLanguage = (lang: unknown): string | undefined =>
  typeof lang === "string" && all[lang] ? lang : undefined;

/**
 * Runs `fn` with the texts for recipients – sections "mail" and "templates", dates and numbers – in `lang` (if the
 * server has it; otherwise LANGUAGE). Only around what the recipients read, never around notes to the sender.
 */
export function inRecipientLanguage<T>(lang: string | undefined, fn: () => T): T {
  const known = knownLanguage(lang);
  return known ? recipientLanguage.run(known, fn) : fn();
}

/** Text of `key` in `lang` (LANGUAGE if unknown) – e.g. the label of a message in the recipients' language. */
export const tIn = (lang: string | undefined, key: string, vars?: Record<string, string | number>) => {
  const file = all[knownLanguage(lang) ?? language] ?? active;
  const dot = key.indexOf(".");
  return format((file as unknown as Record<string, Section>)[key.slice(0, dot)] ?? {}, key.slice(dot + 1), vars);
};

/**
 * Texts of a section for the current context: "server" in the request's language, "mail" and "templates" in the
 * recipients' language inside inRecipientLanguage, all others in LANGUAGE.
 */
function sectionTexts(section: (typeof SECTIONS)[number]): Section {
  if (section === "server") {
    const lang = requestLanguage.getStore();
    if (lang && all[lang]) return all[lang].server;
  }
  if (section === "mail" || section === "templates") {
    const lang = recipientLanguage.getStore();
    if (lang && all[lang]) return all[lang][section];
  }
  return active[section] ?? {};
}

/** BCP 47 locale for dates and numbers ("de-DE", "en-GB") – the recipients' inside inRecipientLanguage. */
export const locale = () => {
  const lang = recipientLanguage.getStore();
  return (lang && all[lang]?._meta.locale) || active._meta.locale;
};

/** Date in the language's format and the configured time zone ("19.11.2026"). */
export const formatDate = (ms: number) => new Date(ms).toLocaleDateString(locale(), { timeZone: config.timeZone });

/** Picks the plural form for `count`, then fills {name} placeholders (unknown ones stay visible). */
export function format(texts: Section, key: string, vars: Record<string, string | number> = {}): string {
  const count = vars.count;
  const form = typeof count === "number" ? (count === 1 ? `${key}_one` : `${key}_other`) : undefined;
  const text = (form !== undefined ? texts[form] : undefined) ?? texts[key] ?? key;
  return text.replace(/\{(\w+)\}/g, (match, name: string) => (name in vars ? String(vars[name]) : match));
}

/**
 * Text: t("server.notConnected"), t("mail.access", { count: 3 }). "server" texts come in the language of the request
 * (withRequestLanguage), everything else in LANGUAGE.
 */
export function t(key: string, vars?: Record<string, string | number>): string {
  const dot = key.indexOf(".");
  const section = key.slice(0, dot) as (typeof SECTIONS)[number];
  return format(sectionTexts(section), key.slice(dot + 1), vars);
}

/** All texts of a section (e.g. the built-in templates) – in the recipients' language inside inRecipientLanguage. */
export const texts = (section: (typeof SECTIONS)[number]): Section => sectionTexts(section);

const escapeMarkup = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Non-ASCII as \uXXXX: the script stays ASCII whatever the runtime assumes as encoding. */
const asciiJs = (s: string) => s.replace(/[^\x00-\x7e]/g, (c) => String.fromCharCode(92) + "u" + c.charCodeAt(0).toString(16).padStart(4, "0"));

/** The marker in the add-in bundles that receives the texts (see addin/src/shared/i18n.ts). */
export const BUNDLE_MARKER = "%%CLOUDSHARE_I18N%%";

const textOf = (file: LocaleFile, section: string, key: string) =>
  format((file as unknown as Record<string, Section>)[section] ?? {}, key);

/** HTML or XML with %%t:section.key%% as the escaped text of `lang`, %%LANG%% / %%LOCALE%% its language / locale. */
function fillMarkup(content: string, languages: Record<string, LocaleFile>, lang: string): string {
  const file = languages[lang];
  return content
    .replace(/%%t:(\w+)\.([\w.]+)%%/g, (_m, section: string, key: string) => escapeMarkup(textOf(file, section, key)))
    .replaceAll("%%LANG%%", lang)
    .replaceAll("%%LOCALE%%", file._meta.locale);
}

/**
 * manifest.xml: an element `<X DefaultValue="%%t:section.key%%"/>` gets the text of `fallback` (English) as
 * DefaultValue and one Override per other language (`<bt:Override>` inside bt:String, `<Override>` at the top level),
 * so Outlook shows the ribbon and the description in the client's language. %%LOCALE%% is the fallback's locale.
 */
function localizeManifest(xml: string, languages: Record<string, LocaleFile>, fallback: string): string {
  const others = Object.keys(languages).filter((lang) => lang !== fallback).sort();
  const withOverrides = xml.replace(
    /<((?:\w+:)?\w+)(\s[^>]*?)DefaultValue="%%t:(\w+)\.([\w.]+)%%"([^>]*?)\s*\/>/g,
    (_m, tag: string, before: string, section: string, key: string, after: string) => {
      const prefix = tag.includes(":") ? `${tag.split(":")[0]}:` : "";
      const value = (lang: string) => escapeMarkup(textOf(languages[lang], section, key));
      const overrides = others
        .map((lang) => `<${prefix}Override Locale="${languages[lang]._meta.locale}" Value="${value(lang)}"/>`)
        .join("");
      return `<${tag}${before}DefaultValue="${value(fallback)}"${after}>${overrides}</${tag}>`;
    },
  );
  return fillMarkup(withOverrides, languages, fallback);
}

/**
 * Puts the languages into the add-in files (run once at startup on the copies in /app/public):
 *   *.js    the marker string becomes a JSON string { fallback, languages: { <lang>: { locale, messages } } } with the
 *           "addin" texts of every language – the add-in picks the client's language (addin/src/shared/i18n.ts)
 *   *.html  with text markers: the page itself in `base` (LANGUAGE, the first guess) and <name>.<lang>.html for
 *           every language, which the taskpane loads when the client has another language
 *   *.xml   the manifest: English as default, Overrides for the other languages
 * Returns the number of changed files (without the language copies).
 */
export function localizePublic(dir: string, languages: Record<string, LocaleFile> = all, base = language): number {
  const names = Object.keys(languages).sort();
  const fallback = languages[FALLBACK_LANGUAGE] ? FALLBACK_LANGUAGE : base;
  const bundle = JSON.stringify({
    fallback,
    languages: Object.fromEntries(
      names.map((lang) => [lang, { locale: languages[lang]._meta.locale, name: languages[lang]._meta.name, messages: languages[lang].addin }]),
    ),
  });
  const jsLiteral = asciiJs(JSON.stringify(bundle));
  let changed = 0;
  for (const name of readdirSync(dir)) {
    const full = path.join(dir, name);
    const ext = path.extname(name);
    if (![".js", ".html", ".xml"].includes(ext)) continue;
    const before = readFileSync(full, "utf8");
    let after: string;
    if (ext === ".js") {
      after = before.replace(/(["'`])%%CLOUDSHARE_I18N%%\1/g, () => jsLiteral);
    } else if (ext === ".xml") {
      after = localizeManifest(before, languages, fallback);
    } else {
      if (/%%t:/.test(before)) {
        for (const lang of names) writeFileSync(path.join(dir, `${name.slice(0, -ext.length)}.${lang}${ext}`), fillMarkup(before, languages, lang));
      }
      after = fillMarkup(before, languages, languages[base] ? base : fallback);
    }
    if (after !== before) {
      writeFileSync(full, after);
      changed++;
    }
  }
  return changed;
}
