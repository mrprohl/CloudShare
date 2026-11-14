/**
 * Language files (locales/*.json) and the language machinery (src/i18n.ts): every file has the same texts and
 * placeholders as English, every key the code uses exists, Outlook's length limits hold in every language, and the
 * add-in files get the texts at startup.
 */
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { runInNewContext } from "node:vm";
import { describe, expect, it } from "vitest";
import { builtinLocalesDir, format, loadLocale, localizePublic, SECTIONS, t, type LocaleFile } from "../src/i18n.js";
import { gsmLength, isGsm, TEMPLATE_FIELDS } from "../src/templates.js";

const repo = path.resolve(import.meta.dirname, "../..");
const languages = readdirSync(builtinLocalesDir).filter((f) => f.endsWith(".json")).map((f) => f.slice(0, -5));
const raw = (lang: string) => JSON.parse(readFileSync(path.join(builtinLocalesDir, `${lang}.json`), "utf8")) as LocaleFile;
const en = raw("en");
const placeholders = (text: string) => [...text.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();

/** All source files below `dir` with one of the extensions. */
function sources(dir: string, exts: string[]): string[] {
  return readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) return sources(full, exts);
    return exts.includes(path.extname(entry.name)) ? [full] : [];
  });
}

/** A key used by the code exists as such or in both plural forms. */
const known = (section: Record<string, string>, key: string) =>
  key in section || (`${key}_one` in section && `${key}_other` in section);

describe("language files", () => {
  it("include German and English", () => {
    expect(languages).toEqual(expect.arrayContaining(["de", "en"]));
  });

  it.each(languages)("%s has the same texts and placeholders as English", (lang) => {
    const file = raw(lang);
    expect(file._meta.name).toBeTruthy();
    expect(file._meta.locale).toMatch(/^[a-z]{2,3}(-[A-Z]{2})?$/);
    for (const section of SECTIONS) {
      expect(Object.keys(file[section]).sort(), `${lang}: keys of "${section}"`).toEqual(Object.keys(en[section]).sort());
      for (const [key, text] of Object.entries(en[section])) {
        expect(placeholders(file[section][key]), `${lang}: placeholders of ${section}.${key}`).toEqual(placeholders(text));
      }
    }
  });

  it.each(languages)("%s stays within Outlook's limits", (lang) => {
    const file = loadLocale(lang);
    expect(file.addin["alert.button"].length).toBeLessThanOrEqual(20); // Smart Alert button
    expect(file.addin["notify.action"].length).toBeLessThanOrEqual(30); // info bar action
    expect(file.manifest.group.length).toBeLessThanOrEqual(125); // ShortString
    expect(file.manifest.tooltip.length).toBeLessThanOrEqual(250); // LongString
    expect(file.manifest.readTooltip.length).toBeLessThanOrEqual(250);
    expect(file.manifest.description.length).toBeLessThanOrEqual(250);
  });

  it.each(languages)("%s has a usable SMS text", (lang) => {
    const sms = loadLocale(lang).templates.sms_template;
    expect(sms).toContain("{password}");
    expect(isGsm(sms)).toBe(true);
    expect(gsmLength(sms.replace("{password}", "x".repeat(12)).replace("{sender_name}", "x".repeat(25)))).toBeLessThanOrEqual(160);
    expect(Object.keys(loadLocale(lang).templates).sort()).toEqual(Object.keys(TEMPLATE_FIELDS).sort());
  });
});

describe("keys used by the code", () => {
  it("server: every t(\"section.key\") exists", () => {
    const missing: string[] = [];
    for (const file of sources(path.join(repo, "server/src"), [".ts"])) {
      for (const m of readFileSync(file, "utf8").matchAll(/"(server|mail|manifest|templates)\.(\w+)"/g)) {
        if (!known((en as unknown as Record<string, Record<string, string>>)[m[1]], m[2])) missing.push(`${path.basename(file)}: ${m[1]}.${m[2]}`);
      }
    }
    expect(missing).toEqual([]);
  });

  it("add-in: every key in the TypeScript sources and every text marker exists", () => {
    const prefixes = new Set(Object.keys(en.addin).map((key) => key.split(".")[0]));
    const missing: string[] = [];
    for (const file of sources(path.join(repo, "addin/src"), [".ts"])) {
      const code = readFileSync(file, "utf8");
      const keys = [
        ...[...code.matchAll(/"([a-z][A-Za-z]*)\.([A-Za-z][\w.]*)"/g)].filter((m) => prefixes.has(m[1])).map((m) => `${m[1]}.${m[2]}`),
        ...[...code.matchAll(/\bt\(\s*"(\w+)"/g)].map((m) => m[1]),
      ];
      for (const key of keys) if (!known(en.addin, key)) missing.push(`${path.basename(file)}: ${key}`);
      // Dynamic SSO keys: t(`sso.${code}`) – every listed code needs a text.
      for (const m of code.matchAll(/SSO_CODES = \[([\d, ]+)\]/g)) {
        for (const ssoCode of m[1].split(",").map((s) => s.trim())) if (!known(en.addin, `sso.${ssoCode}`)) missing.push(`sso.${ssoCode}`);
      }
    }
    for (const file of [...sources(path.join(repo, "addin/src"), [".html"]), path.join(repo, "addin/manifest.xml")]) {
      for (const m of readFileSync(file, "utf8").matchAll(/%%t:(\w+)\.([\w.]+)%%/g)) {
        if (!known((en as unknown as Record<string, Record<string, string>>)[m[1]] ?? {}, m[2])) missing.push(`${path.basename(file)}: ${m[1]}.${m[2]}`);
      }
    }
    expect(missing).toEqual([]);
  });
});

describe("i18n", () => {
  it("uses German by default and picks plural forms", () => {
    expect(t("server.notConnected")).toBe("Nextcloud ist noch nicht verbunden.");
    expect(t("mail.access", { count: 1 })).toBe("nur einmal abrufbar");
    expect(t("mail.access", { count: 5 })).toBe("bis zu 5 Abrufe");
    expect(format({ a: "x {name} {other}" }, "a", { name: "Anna" })).toBe("x Anna {other}");
    expect(t("server.doesNotExist")).toBe("doesNotExist");
  });

  it("adds an own language from the custom directory, English filling the gaps", () => {
    const custom = mkdtempSync(path.join(tmpdir(), "cloudshare-locales-"));
    writeFileSync(
      path.join(custom, "fr.json"),
      JSON.stringify({ _meta: { name: "Français", locale: "fr-FR" }, addin: { "submit.button": "Insérer comme lien" } }),
    );
    const fr = loadLocale("fr", { builtin: builtinLocalesDir, custom });
    expect(fr._meta.locale).toBe("fr-FR");
    expect(fr.addin["submit.button"]).toBe("Insérer comme lien");
    expect(fr.addin["tabs.send"]).toBe(en.addin["tabs.send"]);
  });

  it("changes single texts of a built-in language from the custom directory", () => {
    const custom = mkdtempSync(path.join(tmpdir(), "cloudshare-locales-"));
    writeFileSync(path.join(custom, "de.json"), JSON.stringify({ addin: { "tabs.shares": "Freigaben" } }));
    const de = loadLocale("de", { builtin: builtinLocalesDir, custom });
    expect(de.addin["tabs.shares"]).toBe("Freigaben");
    expect(de.addin["tabs.send"]).toBe("Senden");
    expect(de._meta.locale).toBe("de-DE");
  });

  it("refuses an unknown language with the available ones", () => {
    expect(() => loadLocale("xx", { builtin: builtinLocalesDir, custom: "/nonexistent" })).toThrow(/available: .*de.*en/);
    expect(() => loadLocale("../etc/passwd")).toThrow(/not a language code/);
  });

  it("fills the add-in files at startup", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "cloudshare-public-"));
    mkdirSync(path.join(dir, "assets"));
    writeFileSync(path.join(dir, "taskpane.html"), '<html lang="%%LANG%%"><button title="%%t:addin.shares.refresh%%">%%t:addin.submit.button%%</button></html>');
    writeFileSync(path.join(dir, "manifest.xml"), '<DefaultLocale>%%LOCALE%%</DefaultLocale><Description DefaultValue="%%t:manifest.description%%"/>');
    // Minified bundle: the marker is a plain string literal.
    writeFileSync(path.join(dir, "commands.js"), 'var bundle="%%CLOUDSHARE_I18N%%";result=JSON.parse(bundle);');
    const english = loadLocale("en");
    english.addin["submit.button"] = 'Insert <"link"> & go';
    expect(localizePublic(dir, english, "en")).toBe(3);

    const html = readFileSync(path.join(dir, "taskpane.html"), "utf8");
    expect(html).toBe('<html lang="en"><button title="Refresh">Insert &lt;&quot;link&quot;&gt; &amp; go</button></html>');
    expect(readFileSync(path.join(dir, "manifest.xml"), "utf8")).toContain("<DefaultLocale>en-GB</DefaultLocale>");
    const js = readFileSync(path.join(dir, "commands.js"), "utf8");
    expect(js).not.toContain("%%");
    expect(js).toMatch(/^[\x00-\x7f]*$/); // ASCII only
    const context: { result?: { lang: string; locale: string; messages: Record<string, string> } } = {};
    runInNewContext(js, context);
    expect(context.result?.lang).toBe("en");
    expect(context.result?.messages["alert.button"]).toBe("Open CloudShare");
    expect(context.result?.messages["files.remove"]).toBe("Remove “{name}”");
  });
});
