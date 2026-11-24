import { describe, expect, it } from "vitest";
import { passwordMailHtml, smsProblemMailHtml } from "../src/passwordMail.js";

const base = { sendUrl: "https://vault.example.test/#/send/abc/key", maxAccessCount: 3, subject: "Unterlagen", expireDays: 7 };

describe("password mail", () => {
  it("looks like the link card: one card with badges, button and note", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["Angebot.pdf"] });
    expect(html).toContain("border-left:4px solid #0f6cbd"); // brand accent as in the link card
    expect(html).toContain("Kennwort für Ihre Datei");
    expect(html).toContain("die folgende Datei aus der Nachricht „Unterlagen“");
    expect(html).toMatch(/background:#c4314b[^>]*>PDF</);
    expect(html).toMatch(/<a href="https:\/\/vault\.example\.test\/#\/send\/abc\/key"[^>]*><span[^>]*>Kennwort anzeigen<\/span><\/a>/);
    expect(html).toContain("bis zu 3 Abrufe · läuft nach 7 Tagen ab");
    expect(html).toContain("Falls der Button nicht funktioniert"); // plain link for clients without buttons
    // Fits files, upload links and messages alike – no "Download-Link".
    expect(html).toContain("Geben Sie das Kennwort ein, wenn Sie den Link öffnen.");
  });

  it("uses only inline styles and tables (mail-client safe)", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf", "b.zip"] });
    expect(html).not.toMatch(/class=|<style|<script/);
  });

  it("names every file of the step and uses the plural", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf", "b.zip", "Upload-Link für Ihre Dateien"] });
    expect(html).toContain("Kennwort für 3 Links"); // an upload link among them: "Links", not "Dateien"
    expect(html).toContain("die folgenden Links");
    expect(html).toContain(">ZIP<");
    expect(html).toContain("&#8593;"); // upload link: arrow badge instead of a file type
    expect(html).toContain("Geben Sie das Kennwort ein, wenn Sie einen der Links öffnen.");
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

  it("ends with the sender's signature, before the line about the automatic mail", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf"], signature: "<b>Max Mustermann</b><br>Projektleitung" });
    const signature = html.indexOf("<b>Max Mustermann</b>");
    expect(signature).toBeGreaterThan(html.indexOf("Falls der Button nicht funktioniert"));
    expect(signature).toBeLessThan(html.indexOf("automatisch"));
    expect(passwordMailHtml({ ...base, fileNames: ["a.pdf"] })).not.toContain("Max Mustermann");
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

describe("password mail – limits and the mail after the confirmation", () => {
  it("says in a sentence how often and how long the button works", () => {
    const html = passwordMailHtml({ ...base, fileNames: ["a.pdf"], maxAccessCount: 3, expiresOn: "25.11.2026" });
    expect(html).toContain("Der Button funktioniert 3-mal und bis 25.11.2026 – notieren Sie sich das Kennwort, bis Sie alle Dateien heruntergeladen haben.");
    expect(passwordMailHtml({ ...base, fileNames: ["a.pdf"], maxAccessCount: 1, expiresOn: "25.11.2026" }))
      .toContain("Der Button funktioniert nur einmal und bis 25.11.2026");
    // A confidential message among them – named by its label in any language: "Links".
    expect(passwordMailHtml({ ...base, fileNames: ["a.pdf", "Confidential message"], maxAccessCount: 3, expiresOn: "25.11.2026" }))
      .toContain("bis Sie alle Links geöffnet haben.");
  });

  it("right after the confirmation: two numbered steps in the card, each with its button and plain link", () => {
    const html = passwordMailHtml({
      ...base, fileNames: ["Angebot.zip"], subject: "Unterlagen", senderName: "Anna Muster", maxAccessCount: 3,
      expiresOn: "25.11.2026", trackingId: "20261118-1932-DHLB", sendUrl: "https://vault.example.test/#/send/abc/key",
      afterConfirmation: { kind: "file", url: "https://cloudshare.example.test/r/s/tok", validUntil: "04.12.2026, 23:59 Uhr" },
    });
    expect(html).toContain("Sie haben eben Ihre E-Mail-Adresse bestätigt. Hier ist das Kennwort für „Angebot.zip“ aus der Nachricht „Unterlagen“ von Anna Muster.");
    expect(html).toContain("Kennwort und Download");
    // Step 1 first, then step 2 – the password link never stands below the download.
    const password = html.indexOf("Kennwort anzeigen und notieren");
    const download = html.indexOf("Datei herunterladen – dort das Kennwort eingeben");
    expect(password).toBeGreaterThan(0);
    expect(download).toBeGreaterThan(password);
    // Each step: its button and its address as plain text (the download to the gateway, never to Nextcloud).
    expect(html.match(/https:\/\/vault\.example\.test\/#\/send\/abc\/key/g)).toHaveLength(3); // href of button, href and text of the link
    expect(html.match(/https:\/\/cloudshare\.example\.test\/r\/s\/tok(?!\/)/g)).toHaveLength(3);
    // Under step 1 the way to a new password link (the gateway's address form, also for a confirmed browser).
    expect(html).toContain('href="https://cloudshare.example.test/r/s/tok/password"');
    expect(html.indexOf('href="https://cloudshare.example.test/r/s/tok"')).toBeGreaterThan(download);
    expect(html).toContain(">Zum Download</span></a>");
    // How often and how long the password link works: right under step 1; the case number at the bottom of the card.
    const limits = html.indexOf("Kennwort-Link: 3-mal, bis 25.11.2026");
    expect(limits).toBeGreaterThan(password);
    expect(limits).toBeLessThan(download);
    // Under step 2 until when the link works; then the case number at the bottom.
    const valid = html.indexOf("Gültig bis 04.12.2026, 23:59 Uhr");
    expect(valid).toBeGreaterThan(download);
    expect(html.indexOf("Vorgang 20261118-1932-DHLB")).toBeGreaterThan(valid);
    expect(html).not.toContain("Falls der Button nicht funktioniert");
    expect(html).not.toContain("<ol");
  });

});
