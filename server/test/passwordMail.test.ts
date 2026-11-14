import { describe, expect, it } from "vitest";
import { passwordMailHtml } from "../src/passwordMail.js";

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
