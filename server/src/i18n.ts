/**
 * Languages. Every text users or recipients see comes from a language file: locales/<language>.json in the
 * repository (shipped in the image as /app/locales) and optionally CUSTOM_LOCALES_DIR (default /config/locales,
 * mounted from env/config/locales) for an own language or single changed texts without a new image.
 *
 * LANGUAGE (default "de") picks the file. Missing texts fall back to English, then to the key itself, so a partly
 * translated file still works. Sections of a file:
 *   addin      taskpane, Smart Alerts, info bar, link cards – injected into the add-in files at startup
 *   manifest   texts of manifest.xml (description, ribbon)
 *   server     error messages the taskpane shows
 *   mail       password mail
 *   templates  built-in texts of templates.ts (SMS, password mail, download confirmation)
 *
 * Placeholders are {name}. Plural forms: "key_one" (count 1) and "key_other"; t(key, { count }) picks one.
 */
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

/** Tests only: switch the language. */
export function setLocale(file: LocaleFile): void {
  active = file;
}

/** BCP 47 locale of the active language for dates and numbers ("de-DE", "en-GB"). */
export const locale = () => active._meta.locale;

/** Picks the plural form for `count`, then fills {name} placeholders (unknown ones stay visible). */
export function format(texts: Section, key: string, vars: Record<string, string | number> = {}): string {
  const count = vars.count;
  const form = typeof count === "number" ? (count === 1 ? `${key}_one` : `${key}_other`) : undefined;
  const text = (form !== undefined ? texts[form] : undefined) ?? texts[key] ?? key;
  return text.replace(/\{(\w+)\}/g, (match, name: string) => (name in vars ? String(vars[name]) : match));
}

/** Text of the active language: t("server.notConnected"), t("mail.access", { count: 3 }). */
export function t(key: string, vars?: Record<string, string | number>): string {
  const dot = key.indexOf(".");
  const section = key.slice(0, dot) as (typeof SECTIONS)[number];
  return format(active[section] ?? {}, key.slice(dot + 1), vars);
}

/** All texts of a section (e.g. the built-in templates). */
export const texts = (section: (typeof SECTIONS)[number]): Section => active[section];

const escapeMarkup = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Non-ASCII as \uXXXX: the script stays ASCII whatever the runtime assumes as encoding. */
const asciiJs = (s: string) => s.replace(/[\u007f-￿]/g, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, "0")}`);

/** The marker in the add-in bundles that receives the texts (see addin/src/shared/i18n.ts). */
export const BUNDLE_MARKER = "%%CLOUDSHARE_I18N%%";

/**
 * Puts the active language into the add-in files (run once at startup on the copies in /app/public):
 *   *.js           the marker string becomes a JSON string with lang, locale and the "addin" texts
 *   *.html / *.xml %%t:section.key%% becomes the escaped text, %%LANG%% / %%LOCALE%% the language / locale
 * Returns the number of changed files.
 */
export function localizePublic(dir: string, file: LocaleFile = active, lang = language): number {
  const bundle = JSON.stringify({ lang, locale: file._meta.locale, messages: file.addin });
  const jsLiteral = asciiJs(JSON.stringify(bundle));
  let changed = 0;
  for (const name of readdirSync(dir)) {
    const full = path.join(dir, name);
    const ext = path.extname(name);
    if (![".js", ".html", ".xml"].includes(ext)) continue;
    const before = readFileSync(full, "utf8");
    const after =
      ext === ".js"
        ? before.replace(/(["'`])%%CLOUDSHARE_I18N%%\1/g, () => jsLiteral)
        : before
            .replace(/%%t:(\w+)\.([\w.]+)%%/g, (_m, section: string, key: string) =>
              escapeMarkup(format((file as unknown as Record<string, Section>)[section] ?? {}, key)),
            )
            .replaceAll("%%LANG%%", lang)
            .replaceAll("%%LOCALE%%", file._meta.locale);
    if (after !== before) {
      writeFileSync(full, after);
      changed++;
    }
  }
  return changed;
}
