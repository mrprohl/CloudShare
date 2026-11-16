// Test tooling only (UI tests, manifest validation, bundle test): fills the language markers of the add-in files like
// the server does at container start (server/src/i18n.ts, localizePublic). English is the base of every language.
import { readdirSync, readFileSync } from "node:fs";

const read = (language) =>
  JSON.parse(readFileSync(new URL(`../../locales/${language}.json`, import.meta.url), "utf8"));

export function loadLocale(language = "de") {
  const base = read("en");
  const own = read(language);
  const merged = { _meta: { ...base._meta, ...own._meta } };
  for (const section of ["addin", "manifest", "server", "mail", "templates"]) {
    merged[section] = { ...base[section], ...own[section] };
  }
  return merged;
}

/** Every language of locales/ – { de: …, en: … }. */
export function loadLanguages() {
  const names = readdirSync(new URL("../../locales/", import.meta.url))
    .filter((file) => file.endsWith(".json"))
    .map((file) => file.slice(0, -5));
  return Object.fromEntries(names.map((name) => [name, loadLocale(name)]));
}

const FALLBACK = "en";
const escapeMarkup = (s) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const asciiJs = (s) => s.replace(/[^\x00-\x7e]/g, (c) => String.fromCharCode(92) + "u" + c.charCodeAt(0).toString(16).padStart(4, "0"));
const textOf = (file, section, key) => file[section]?.[key] ?? key;

function fillMarkup(content, languages, language) {
  const file = languages[language];
  return content
    .replace(/%%t:(\w+)\.([\w.]+)%%/g, (_m, section, key) => escapeMarkup(textOf(file, section, key)))
    .replaceAll("%%LANG%%", language)
    .replaceAll("%%LOCALE%%", file._meta.locale);
}

// Own locale and regional variants (_meta.regions) – as server/src/i18n.ts.
const manifestLocales = (meta) =>
  [...new Set([meta.locale, ...(Array.isArray(meta.regions) ? meta.regions : [])])].filter((l) => /^[a-z]{2,3}-[A-Za-z]{2,4}$/.test(l));

function localizeManifest(xml, languages) {
  const others = Object.keys(languages).filter((lang) => lang !== FALLBACK).sort();
  const withOverrides = xml.replace(
    /<((?:\w+:)?\w+)(\s[^>]*?)DefaultValue="%%t:(\w+)\.([\w.]+)%%"([^>]*?)\s*\/>/g,
    (_m, tag, before, section, key, after) => {
      const prefix = tag.includes(":") ? `${tag.split(":")[0]}:` : "";
      const value = (lang) => escapeMarkup(textOf(languages[lang], section, key));
      const overrides = others
        .map((lang) => manifestLocales(languages[lang]._meta).map((locale) => `<${prefix}Override Locale="${locale}" Value="${value(lang)}"/>`).join(""))
        .join("");
      return `<${tag}${before}DefaultValue="${value(FALLBACK)}"${after}>${overrides}</${tag}>`;
    },
  );
  return fillMarkup(withOverrides, languages, FALLBACK);
}

/**
 * Content of an add-in file (by extension) with the texts filled in: a bundle gets every language, a page the texts
 * of `language`, the manifest English with Overrides for the other languages.
 */
export function localize(content, ext, language = "de", languages = loadLanguages()) {
  if (ext === ".js") {
    const bundle = JSON.stringify({
      fallback: FALLBACK,
      languages: Object.fromEntries(
        Object.keys(languages)
          .sort()
          .map((lang) => [lang, { locale: languages[lang]._meta.locale, name: languages[lang]._meta.name, messages: languages[lang].addin }]),
      ),
    });
    return content.replace(/(["'`])%%CLOUDSHARE_I18N%%\1/g, () => asciiJs(JSON.stringify(bundle)));
  }
  if (ext === ".xml") return localizeManifest(content, languages);
  return fillMarkup(content, languages, language);
}
