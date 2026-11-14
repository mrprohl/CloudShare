// Test tooling only (UI tests, manifest validation): fills the language markers of the add-in files like the server
// does at container start (server/src/i18n.ts, localizePublic). English is the base, the language file on top.
import { readFileSync } from "node:fs";

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

const escapeMarkup = (s) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
const asciiJs = (s) => s.replace(/[\u007f-￿]/g, (c) => `\u${c.charCodeAt(0).toString(16).padStart(4, "0")}`);

/** Content of an add-in file (by extension) with the texts of `locale` (from loadLocale) filled in. */
export function localize(content, ext, locale, language = "de") {
  if (ext === ".js") {
    const bundle = JSON.stringify({ lang: language, locale: locale._meta.locale, messages: locale.addin });
    return content.replace(/(["'`])%%CLOUDSHARE_I18N%%\1/g, () => asciiJs(JSON.stringify(bundle)));
  }
  return content
    .replace(/%%t:(\w+)\.([\w.]+)%%/g, (_m, section, key) => escapeMarkup(locale[section]?.[key] ?? key))
    .replaceAll("%%LANG%%", language)
    .replaceAll("%%LOCALE%%", locale._meta.locale);
}
