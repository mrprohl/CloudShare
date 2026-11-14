import { describe, expect, it } from "vitest";
import { passwordMailHtml, smsProblemMailHtml } from "../src/passwordMail.js";

const base = { sendUrl: "https://vault.example.test/#/send/abc/key", maxAccessCount: 3, subject: "Unterlagen", expireDays: 7 };

describe("password mail", () => {
  it("looks like the link card: one card with badges, button and note", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["Angebot.pdf"] });
    expect(html).toContain("border-left:4px solid #0f6cbd"); // brand accent as in the link card
    expect(html).toContain("Kennwort für Ihre Freigabe");
    expect(html).toContain("die folgende Freigabe aus der Nachricht „Unterlagen“");
    expect(html).toMatch(/background:#c4314b[^>]*>PDF</);
    expect(html).toMatch(/<a href="https:\/\/vault\.example\.test\/#\/send\/abc\/key"[^>]*>Kennwort anzeigen<\/a>/);
    expect(html).toContain("bis zu 3 Abrufe · läuft nach 7 Tagen ab");
    expect(html).toContain("Falls der Button nicht funktioniert"); // plain link for clients without buttons
  });

  it("uses only inline styles and tables (mail-client safe)", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf", "b.zip"] });
    expect(html).not.toMatch(/class=|<style|<script/);
  });

  it("names every file of the step and uses the plural", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf", "b.zip", "Upload-Link für Ihre Dateien"] });
    expect(html).toContain("Kennwort für 3 Freigaben");
    expect(html).toContain("die folgenden Freigaben");
    expect(html).toContain(">ZIP<");
    expect(html).toContain("&#8593;"); // upload link: arrow badge instead of a file type
  });

  it("escapes file names, subject and link", () => {
    const html = passwordMailHtml({
      ...base,
      fileNames: ['<img src=x onerror="a()">.pdf'],
      subject: "<b>Betreff</b>",
      sendUrl: 'https://x.test/"><script>',
    });
    expect(html).not.toContain("<img");
    expect(html).not.toContain("<b>Betreff");
    expect(html).not.toContain('"><script>');
  });

  it("states a single access", () => {
    expect(passwordMailHtml({ ...base, maxAccessCount: 1, fileNames: ["a.pdf"] })).toContain("nur einmal abrufbar");
  });
});

describe("password mail for a confidential message", () => {
  it("shows an envelope instead of a file type", () => {
    const html = passwordMailHtml({ fileNames: ["Vertrauliche Nachricht"], sendUrl: "https://vault.example.test/#/send/p", subject: "Zugang", expireDays: 7 });
    expect(html).toContain("&#9993;");
    expect(html).toContain("Vertrauliche Nachricht");
  });
});

describe("expiry of the password link", () => {
  it("names the last day when the step knows it, otherwise the number of days", () => {
    const html = passwordMailHtml({ fileNames: ["a.pdf"], sendUrl: "https://vault.example.test/#/send/p", subject: "", expireDays: 5, expiresOn: "19.11.2026" });
    expect(html).toContain("läuft am 19.11.2026 ab");
    expect(passwordMailHtml({ fileNames: ["a.pdf"], sendUrl: "https://vault.example.test/#/send/p", subject: "", expireDays: 5 })).toContain("läuft nach 5 Tagen ab");
  });
});

describe("common layout of all automatic mails", () => {
  it("ends every mail with the line that CloudShare sent it", () => {
    expect(passwordMailHtml({ ...base, fileNames: ["a.pdf"] })).toContain("Diese Nachricht wurde automatisch von CloudShare versendet.");
    expect(smsProblemMailHtml({ fileNames: ["a.pdf"], subject: "x", to: "+49 170 •••• 67" })).toContain("automatisch von CloudShare");
  });

  it("notes a failed SMS in a red card with files, number and reason – never with a password or link", () => {
    const html = smsProblemMailHtml({ fileNames: ["Angebot.pdf", "Plan.zip"], subject: "Unterlagen", to: "+49 170 •••• 67", reason: "kein Guthaben" });
    expect(html).toContain("border-left:4px solid #c50f1f");
    expect(html).toContain("SMS mit dem Kennwort nicht versendet");
    expect(html).toContain("die SMS mit dem Kennwort für Ihre Nachricht „Unterlagen“ konnte nicht versendet werden.");
    expect(html).toMatch(/background:#c4314b[^>]*>PDF</);
    expect(html).toContain("an +49 170 •••• 67 · Grund: kein Guthaben");
    expect(html).toContain("„Kennwort“"); // how to send it again
    expect(html).not.toMatch(/<a |class=|<style|<script/);
  });

  it("says when seven.io reported the SMS as not delivered later", () => {
    const html = smsProblemMailHtml({ fileNames: ["a.pdf"], subject: "", to: "+49 170 •••• 67" });
    expect(html).toContain("SMS mit dem Kennwort nicht zugestellt");
    expect(html).toContain("„(ohne Betreff)“ wurde nicht zugestellt.");
    expect(html).not.toContain("Grund:");
  });
});
