import { describe, expect, it } from "vitest";
import { filesCard, filesText, messageCard, messageText, requestCard, requestText } from "../src/taskpane/mailCard";

const formatSize = (bytes: number) => `${bytes} B`;
const files = [
  { name: "Angebot.pdf", size: 1000, url: "https://cloud.example.test/s/a" },
  { name: "Daten.XLSX", size: 2000, url: "https://cloud.example.test/s/b" },
  { name: "ohne-endung", size: 3000, url: "https://cloud.example.test/s/c" },
];

describe("filesCard", () => {
  const html = filesCard(files, "27.11.2026", formatSize);

  it("has one download button per file and a plural title", () => {
    expect(html).toContain("3 Dateien wurden für Sie bereitgestellt");
    expect(html.match(/>Herunterladen</g)).toHaveLength(3);
    for (const f of files) expect(html).toContain(`href="${f.url}"`);
  });

  it("shows coloured type badges, also for upper-case and missing extensions", () => {
    expect(html).toContain("background:#c4314b"); // pdf
    expect(html).toMatch(/background:#107c41[^>]*>XLSX</); // xlsx, case-insensitive
    expect(html).toContain(">DATEI<"); // no extension → the "file" badge of the language
  });

  it("shows expiry and password note, no signature (that is in the password mail)", () => {
    expect(html).toContain("Gültig bis 27.11.2026");
    expect(html).toContain("Das Kennwort erhalten Sie in einer separaten Nachricht.");
    expect(html).not.toContain("padding-top:8px");
  });

  it("leaves out the password note for links without password", () => {
    const open = filesCard(files, "27.11.2026", formatSize, false);
    expect(open).toContain("Gültig bis 27.11.2026");
    expect(open).not.toContain("Kennwort");
    expect(filesText(files, "27.11.2026", formatSize, false)).not.toContain("Kennwort");
  });

  it("leaves out an unknown size (links inserted again from 'Verlauf')", () => {
    const file = { name: "Angebot.pdf", size: null, url: "https://cloud.example.test/s/a" };
    expect(filesCard([file], "27.11.2026", formatSize)).not.toContain(" B<");
    expect(filesText([file], "27.11.2026", formatSize)).toContain("- Angebot.pdf: https://cloud.example.test/s/a");
  });

  it("uses only inline styles and tables (mail-client safe)", () => {
    expect(html).not.toMatch(/class=|<style|<script/);
    expect(html.startsWith("<table")).toBe(true);
  });

  it("uses the singular title for one file", () => {
    expect(filesCard([files[0]], "27.11.2026", formatSize)).toContain("Eine Datei wurde für Sie bereitgestellt");
  });

  it("escapes names and URLs", () => {
    const evil = filesCard([{ name: '<img src=x onerror="a()">.pdf', size: 1, url: 'https://x.test/"><script>' }], "27.11.2026", formatSize);
    expect(evil).not.toContain("<img");
    expect(evil).not.toContain("<script>");
    expect(evil).toContain("&lt;img src=x onerror=&quot;a()&quot;&gt;.pdf");
  });
});

describe("requestCard", () => {
  it("contains the upload button and the password note", () => {
    const html = requestCard("https://cloud.example.test/s/up", "27.11.2026");
    expect(html).toContain("Dateien sicher hochladen");
    expect(html).toContain('href="https://cloud.example.test/s/up"');
    expect(html).toContain(">Dateien hochladen<");
    expect(html).toContain("Das Kennwort erhalten Sie in einer separaten Nachricht.");
  });
});

describe("plain-text versions", () => {
  it("list every link, without a signature", () => {
    const text = filesText(files, "27.11.2026", formatSize);
    expect(text).toContain("- Angebot.pdf (1000 B): https://cloud.example.test/s/a");
    expect(text).not.toContain("--\n");
    expect(requestText("https://cloud.example.test/s/up", "27.11.2026")).toContain("https://cloud.example.test/s/up");
  });
});

describe("messageCard", () => {
  it("links the confidential message without its text and points to the password", () => {
    const url = "https://vault.example.test/#/send/abc/key123";
    const html = messageCard(url, "27.11.2026");
    expect(html).toContain("Vertrauliche Nachricht für Sie");
    expect(html).toContain(`href="${url}"`);
    expect(html).toContain(">Nachricht öffnen<");
    expect(html).toContain("Gültig bis 27.11.2026 · Das Kennwort erhalten Sie in einer separaten Nachricht.");
    expect(messageText(url, "27.11.2026")).toContain(`Vertrauliche Nachricht für Sie (gültig bis 27.11.2026):\n${url}`);
  });
});

describe("one layout for every card", () => {
  it("puts the button first and the file below it in grey; buttons are filled cells (classic Outlook)", () => {
    const html = filesCard(files, "27.11.2026", formatSize);
    expect(html).toContain("Über die folgenden Links laden Sie die Dateien herunter.");
    // Each block: the button, then type, name and size.
    expect(html.indexOf('href="https://cloud.example.test/s/a"')).toBeLessThan(html.indexOf("Angebot.pdf · 1000 B"));
    expect(html.indexOf("Angebot.pdf · 1000 B")).toBeLessThan(html.indexOf('href="https://cloud.example.test/s/b"'));
    for (const card of [html, messageCard("https://vault.example.test/#/send/m/k", "27.11.2026"), requestCard("https://cloud.example.test/s/up", "27.11.2026")]) {
      expect(card).toContain('bgcolor="#0f6cbd"');
      expect(card).toContain("mso-padding-alt:6px 14px");
      expect(card).not.toMatch(/align="right"/);
    }
  });
});

describe("spacing in classic Outlook", () => {
  it("keeps the space between the file blocks on table cells (Word ignores margins of <div>)", () => {
    const html = filesCard(files, "27.11.2026", formatSize);
    expect(html).toContain('<td style="padding:0px 0 12px 0;">'); // first block: space below
    expect(html).toContain('<td style="padding:12px 0 12px 0;border-top:1px solid #f0f0f0;">'); // middle: line and space
    expect(html).toContain('<td style="padding:12px 0 0px 0;border-top:1px solid #f0f0f0;">'); // last
    expect(html).not.toMatch(/<div style="[^"]*margin:12px/);
  });
});

describe("delivery proof", () => {
  const proof = { id: "20261124-1034-WRB4", gateway: true };
  it("facts in one line, then one sentence in the order things happen – the address confirmed only once", () => {
    const html = filesCard(files, "08.12.2026", formatSize, true, proof);
    expect(html).toContain("Gültig bis 08.12.2026 · Vorgang 20261124-1034-WRB4<br>" +
      "Vor dem Öffnen bestätigen Sie Ihre E-Mail-Adresse mit einem Code, den Sie per E-Mail erhalten – danach erhalten Sie " +
      "das Kennwort. Abrufe werden protokolliert.");
    expect(html.match(/bestätigen Sie/g)).toHaveLength(1);
    expect(html).not.toContain("separaten Nachricht");
  });

  it("without a password: no word about it; the plain text has the same lines", () => {
    expect(filesCard(files, "08.12.2026", formatSize, false, proof)).toContain(
      "Gültig bis 08.12.2026 · Vorgang 20261124-1034-WRB4<br>Vor dem Öffnen bestätigen Sie Ihre E-Mail-Adresse mit einem " +
      "Code, den Sie per E-Mail erhalten. Abrufe werden protokolliert.");
    const text = filesText(files, "08.12.2026", formatSize, true, proof);
    expect(text).toContain("Vorgang 20261124-1034-WRB4\nVor dem Öffnen bestätigen Sie");
  });

  it("a confidential message (no confirmation page): the separate password message, then that opening is recorded", () => {
    expect(messageCard("https://vault.example.com/#/send/x", "08.12.2026", { id: "20261124-1034-WRB4", gateway: false })).toContain(
      "Gültig bis 08.12.2026 · Vorgang 20261124-1034-WRB4<br>Das Kennwort erhalten Sie in einer separaten Nachricht. " +
      "Das Öffnen wird protokolliert.");
  });
});
