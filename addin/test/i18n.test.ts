/**
 * Language of the add-in: it follows Office.context.displayLanguage, a regional variant uses its language, anything
 * else falls back to English (the server writes every language into the bundle, see server/src/i18n.ts).
 */
import { afterEach, describe, expect, it } from "vitest";
import de from "../../locales/de.json";
import en from "../../locales/en.json";
import { forceLanguage, language, locale, pickLanguage, setLanguages, setMessages, t } from "../src/shared/i18n";

const languages = {
  de: { locale: "de-DE", messages: (de as { addin: Record<string, string> }).addin },
  en: { locale: "en-GB", messages: (en as { addin: Record<string, string> }).addin },
};

const withClient = (displayLanguage: string | undefined) => {
  (globalThis as Record<string, unknown>).Office = { context: { displayLanguage } };
};

describe("language of the add-in", () => {
  afterEach(() => {
    delete (globalThis as Record<string, unknown>).Office;
    setMessages(languages.de.messages, "de-DE"); // as test/setup.ts
  });

  it("picks the exact language, the one without region, else English", () => {
    setLanguages(languages);
    expect(pickLanguage("de-DE")).toBe("de");
    expect(pickLanguage("de-CH")).toBe("de");
    expect(pickLanguage("EN-US")).toBe("en");
    expect(pickLanguage("fr-FR")).toBe("en");
    expect(pickLanguage("")).toBe("en");
    expect(pickLanguage("pt-BR", ["de", "en", "pt-br"])).toBe("pt-br");
  });

  it("follows Outlook's display language", () => {
    setLanguages(languages);
    withClient("en-US");
    expect(language()).toBe("en");
    expect(t("tabs.send")).toBe("Send");
    expect(locale()).toBe("en-GB");

    withClient("de-AT");
    expect(language()).toBe("de");
    expect(t("tabs.send")).toBe("Senden");
    expect(locale()).toBe("de-DE");

    withClient("ja-JP");
    expect(t("tabs.send")).toBe("Send");
  });

  it("stays in a forced language (the taskpane page could not be switched)", () => {
    setLanguages(languages);
    withClient("en-US");
    forceLanguage("de");
    expect(t("tabs.send")).toBe("Senden");
    forceLanguage("xx"); // unknown: ignored
    expect(language()).toBe("de");
  });
});
