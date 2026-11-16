import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { defaultTemplates, gsmLength, isGsm, loadTemplates, render, renderHtml, smsFiles, templatesFor, toGsm } from "../src/templates.js";
import { inRecipientLanguage } from "../src/i18n.js";

const dir = mkdtempSync(path.join(tmpdir(), "cloudshare-templates-"));
function file(content: unknown) {
  const target = path.join(dir, `t-${Math.random().toString(36).slice(2)}.json`);
  writeFileSync(target, typeof content === "string" ? content : JSON.stringify(content));
  return target;
}

describe("templates", () => {
  it("uses the built-in texts without a file", () => {
    const warnings: string[] = [];
    const t = templatesFor({ upn: "a@b.de" }, loadTemplates(path.join(dir, "missing.json"), (m) => warnings.push(m)));
    expect(t).toEqual(defaultTemplates());
    expect(warnings).toEqual([]);
    expect(isGsm(defaultTemplates().sms_template)).toBe(true);
  });

  it("resolves user → department → default per field", () => {
    const loaded = loadTemplates(
      file({
        templates: {
          default: { sms_template: "Standard {password}", download_notice_subject: "Default-Betreff" },
          groups: { Ausbildungsberatung: { sms_template: "Gruppe {password}" } },
          users: { "Max.Mustermann@Example.test": { download_notice_subject: "Nur für mich" } },
        },
      }),
      () => undefined,
    );
    const user = templatesFor({ upn: "max.mustermann@example.test", department: "ausbildungsberatung" }, loaded);
    expect(user.sms_template).toBe("Gruppe {password}"); // group beats default
    expect(user.download_notice_subject).toBe("Nur für mich"); // user beats default
    expect(user.mail_password_subject).toBe(defaultTemplates().mail_password_subject); // not set anywhere
    expect(templatesFor({ upn: "x@example.de", department: "Andere" }, loaded).sms_template).toBe("Standard {password}");
  });

  it("uses the texts of the mail's language where a level has them", () => {
    const warnings: string[] = [];
    const loaded = loadTemplates(
      file({
        templates: {
          default: {
            sms_template: "Kennwort:\n{password}",
            languages: { EN: { sms_template: "Password:\n{password}" }, fr: { mail_password_subject: "Mot de passe {password}" } },
          },
          users: { "max@example.test": { download_notice_subject: "Nur für mich", languages: { en: { download_notice_subject: "Just for me" } } } },
        },
      }),
      (m) => warnings.push(m),
    );
    const de = templatesFor({ upn: "max@example.test" }, loaded); // LANGUAGE (de in the tests)
    expect(de.sms_template).toBe("Kennwort:\n{password}");
    expect(de.download_notice_subject).toBe("Nur für mich");
    const en = inRecipientLanguage("en", () => templatesFor({ upn: "max@example.test" }, loaded));
    expect(en.sms_template).toBe("Password:\n{password}");
    expect(en.download_notice_subject).toBe("Just for me");
    expect(en.mail_password_subject).toBe(inRecipientLanguage("en", () => defaultTemplates().mail_password_subject)); // built-in English
    // Checked like any other text: {password} outside the SMS is dropped.
    expect(warnings.some((w) => w.includes("default.languages.fr.mail_password_subject must not contain {password}"))).toBe(true);
  });

  it("drops fields that would leak something and warns about the rest", () => {
    const warnings: string[] = [];
    const loaded = loadTemplates(
      file({
        templates: {
          default: {
            mail_password_intro: "Ihr Kennwort: {password}", // never in a mail
            sms_template: "Datei {filename}: {password}", // never the file in the SMS
            download_notice_body: "Hallo {wer}", // unknown placeholder – kept, warned
            foo: "bar",
          },
          groups: { X: { sms_template: "kein Platzhalter" } }, // SMS without {password}
        },
      }),
      (m) => warnings.push(m),
    );
    const t = templatesFor({ department: "x" }, loaded);
    expect(t.mail_password_intro).toBe(defaultTemplates().mail_password_intro);
    expect(t.sms_template).toBe(defaultTemplates().sms_template);
    expect(t.download_notice_body).toBe("Hallo {wer}");
    expect(warnings.join("\n")).toMatch(/must not contain \{password\}/);
    expect(warnings.join("\n")).toMatch(/must not contain the link or a file name/);
    expect(warnings.join("\n")).toMatch(/unknown placeholders \{wer\}/);
    expect(warnings.join("\n")).toMatch(/unknown field "foo"/);
    expect(warnings.join("\n")).toMatch(/needs \{password\}/);
  });

  it("warns about SMS texts that are too long or not GSM", () => {
    const warnings: string[] = [];
    loadTemplates(file({ templates: { default: { sms_template: `${"x".repeat(150)} {password}` } } }), (m) => warnings.push(m));
    loadTemplates(file({ templates: { default: { sms_template: "Kennwort 🔑 {password}" } } }), (m) => warnings.push(m));
    expect(warnings.join("\n")).toMatch(/billed as several SMS/);
    expect(warnings.join("\n")).toMatch(/non-GSM characters/);
  });

  it("warns if the password is not on a line of its own in the SMS", () => {
    const warnings: string[] = [];
    loadTemplates(file({ templates: { default: { sms_template: "Ihr Kennwort: {password}" } } }), (m) => warnings.push(m));
    expect(warnings.join("\n")).toMatch(/on a line of its own/);
    warnings.length = 0;
    loadTemplates(file({ templates: { default: { sms_template: "Ihr Kennwort:\n{password}" } } }), (m) => warnings.push(m));
    loadTemplates(file({ templates: { default: { sms_template: "{password} ist Ihr Kennwort" } } }), (m) => warnings.push(m));
    expect(warnings).toEqual([]);
  });

  it("falls back to the built-in texts for invalid JSON", () => {
    const warnings: string[] = [];
    expect(templatesFor({}, loadTemplates(file("{ kaputt"), (m) => warnings.push(m)))).toEqual(defaultTemplates());
    expect(warnings[0]).toMatch(/not valid JSON/);
  });

  it("renders placeholders, keeps unknown ones visible and escapes HTML", () => {
    expect(render("Hallo {sender_name}, {unbekannt}", { sender_name: "Max" })).toBe("Hallo Max, {unbekannt}");
    const html = renderHtml("Guten Tag,\n\n<b>{subject}</b>\nZeile", { subject: "<script>" }, "x");
    expect(html).toBe('<p style="x">Guten Tag,</p><p style="x">&lt;b&gt;&lt;script&gt;&lt;/b&gt;<br>Zeile</p>');
    expect(gsmLength("€uro")).toBe(5); // € counts twice
    expect(isGsm("Grüße äöü ß")).toBe(true);
  });

  it("names the files in the SMS: GSM characters, shortened, how many more", () => {
    expect(smsFiles([])).toBe("");
    expect(smsFiles(["Angebot.pdf"])).toBe('"Angebot.pdf"');
    expect(smsFiles(["a.pdf", "b.pdf", "c.pdf"])).toBe('"a.pdf" und 2 weitere');
    expect(toGsm("„Crêpe“ – Señor ł.pdf")).toBe('"Crepe" - Señor _.pdf');
    const long = smsFiles(["Angebot_Neubauprojekt_Rathausplatz_2026_final.pdf"]);
    expect(long).toBe('"Angebot_Neubauproje..._final.pdf"');
    expect(isGsm(smsFiles(["Präsentation – Entwurf „neu“.pptx", "x"]))).toBe(true);
    // The built-in SMS stays one SMS with a long name and sender.
    const sms = render(defaultTemplates().sms_template, { password: "x".repeat(12), sender_name: "x".repeat(25), files: long });
    expect(gsmLength(sms)).toBeLessThanOrEqual(160);
    expect(sms).toContain('Kennwort fuer "Angebot_Neubauproje..._final.pdf" von ');
  });

  it("accepts {files} in an own SMS text", () => {
    const warnings: string[] = [];
    const loaded = loadTemplates(file({ templates: { default: { sms_template: "Code fuer {files}:\n{password}" } } }), (m) => warnings.push(m));
    expect(templatesFor({}, loaded).sms_template).toBe("Code fuer {files}:\n{password}");
    expect(warnings).toEqual([]);
  });
});
