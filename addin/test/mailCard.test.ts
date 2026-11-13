import { describe, expect, it } from "vitest";
import { filesCard, filesText, requestCard, requestText } from "../src/taskpane/mailCard";

const formatSize = (bytes: number) => `${bytes} B`;
const files = [
  { name: "Angebot.pdf", size: 1000, url: "https://cloud.example.test/s/a" },
  { name: "Daten.XLSX", size: 2000, url: "https://cloud.example.test/s/b" },
  { name: "ohne-endung", size: 3000, url: "https://cloud.example.test/s/c" },
];

describe("filesCard", () => {
  const html = filesCard(files, "27.11.2026", "Max Mustermann\nMuster GmbH", formatSize);

  it("has one download button per file and a plural title", () => {
    expect(html).toContain("3 Dateien für Sie bereitgestellt");
    expect(html.match(/>Herunterladen</g)).toHaveLength(3);
    for (const f of files) expect(html).toContain(`href="${f.url}"`);
  });

  it("shows coloured type badges, also for upper-case and missing extensions", () => {
    expect(html).toContain("background:#c4314b"); // pdf
    expect(html).toMatch(/background:#107c41[^>]*>XLSX</); // xlsx, case-insensitive
    expect(html).toContain(">DATE<"); // no extension → "DATEI" shortened to 4 characters
  });

  it("shows expiry, password note and a multi-line footer", () => {
    expect(html).toContain("Gültig bis 27.11.2026");
    expect(html).toContain("Das Kennwort erhalten Sie in einer separaten Nachricht.");
    expect(html).toContain("Max Mustermann<br>Muster GmbH");
  });

  it("uses only inline styles and tables (mail-client safe)", () => {
    expect(html).not.toMatch(/class=|<style|<script/);
    expect(html.startsWith("<table")).toBe(true);
  });

  it("uses the singular title for one file and omits an empty footer", () => {
    const one = filesCard([files[0]], "27.11.2026", "", formatSize);
    expect(one).toContain("Eine Datei für Sie bereitgestellt");
    expect(one).not.toContain("border-top:1px solid #f0f0f0;padding-top:8px");
  });

  it("escapes names, URLs and footer", () => {
    const evil = filesCard(
      [{ name: '<img src=x onerror="a()">.pdf', size: 1, url: 'https://x.test/"><script>' }],
      "27.11.2026",
      "<b>Fett</b>",
      formatSize,
    );
    expect(evil).not.toContain("<img");
    expect(evil).not.toContain("<script>");
    expect(evil).not.toContain("<b>Fett</b>");
    expect(evil).toContain("&lt;img src=x onerror=&quot;a()&quot;&gt;.pdf");
  });
});

describe("requestCard", () => {
  it("contains the upload button and the password note", () => {
    const html = requestCard("https://cloud.example.test/s/up", "27.11.2026", "");
    expect(html).toContain("Dateien sicher hochladen");
    expect(html).toContain('href="https://cloud.example.test/s/up"');
    expect(html).toContain(">Dateien hochladen<");
    expect(html).toContain("Das Kennwort erhalten Sie in einer separaten Nachricht.");
  });
});

describe("plain-text versions", () => {
  it("list every link and include the footer", () => {
    const text = filesText(files, "27.11.2026", "Max Mustermann", formatSize);
    expect(text).toContain("- Angebot.pdf (1000 B): https://cloud.example.test/s/a");
    expect(text).toContain("--\nMax Mustermann");
    expect(requestText("https://cloud.example.test/s/up", "27.11.2026", "")).toContain("https://cloud.example.test/s/up");
  });
});
