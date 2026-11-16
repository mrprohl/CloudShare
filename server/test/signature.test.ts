/**
 * Signature of the password mails (src/signature.ts): plain text with profile placeholders as MAIL_FOOTER always had
 * them, paragraphs separated by an empty line, everything escaped, in the style of the mail text.
 */
import { describe, expect, it } from "vitest";
import { TEXT_STYLE } from "../src/mailLayout.js";
import { fillTemplate, signatureHtml, signatureParagraphs } from "../src/signature.js";

const profile = {
  displayName: "Max Mustermann",
  jobTitle: "Projektleitung",
  company: "Muster GmbH",
  email: "max@example.test",
  phone: "+49 30 1234567",
  department: "",
};
const TEMPLATE = "{displayName} · {jobTitle}\\n{company} · Tel. {phone} · {email}";

describe("fillTemplate (one paragraph)", () => {
  it("fills placeholders and turns \\n into line breaks", () => {
    expect(fillTemplate(TEMPLATE, profile)).toBe("Max Mustermann · Projektleitung\nMuster GmbH · Tel. +49 30 1234567 · max@example.test");
  });

  it("drops a segment with its fixed text when its placeholders are empty", () => {
    expect(fillTemplate(TEMPLATE, { ...profile, phone: "" })).toBe("Max Mustermann · Projektleitung\nMuster GmbH · max@example.test");
  });

  it("drops lines that end up empty, keeps fixed text, treats unknown placeholders as empty", () => {
    expect(fillTemplate("{department}\\n{displayName}", profile)).toBe("Max Mustermann");
    expect(fillTemplate("Vertraulich · {department}", profile)).toBe("Vertraulich");
    expect(fillTemplate("{displayName} · {gibtsNicht}", profile)).toBe("Max Mustermann");
    expect(fillTemplate("{displayName}\r\n{company}", profile)).toBe("Max Mustermann\nMuster GmbH");
  });
});

describe("paragraphs", () => {
  it("are separated by an empty line – typed, with spaces in it, or written as \\n\\n in the env file", () => {
    expect(signatureParagraphs("Mit freundlichen Grüßen\n\n{displayName}\n{jobTitle}", profile)).toEqual([
      "Mit freundlichen Grüßen",
      "Max Mustermann\nProjektleitung",
    ]);
    expect(signatureParagraphs("A\r\n  \r\nB", profile)).toEqual(["A", "B"]);
    expect(signatureParagraphs("{displayName}\\n\\n{company}", profile)).toEqual(["Max Mustermann", "Muster GmbH"]);
  });

  it("drops paragraphs whose placeholders are all empty", () => {
    expect(signatureParagraphs("{displayName}\n\n{department}\n\nEnde", profile)).toEqual(["Max Mustermann", "Ende"]);
  });
});

describe("signatureHtml", () => {
  it("shows each paragraph like the mail text, lines as line breaks", () => {
    expect(signatureHtml({ format: "text", content: "Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone}" }, profile)).toBe(
      `<p style="${TEXT_STYLE}">Mit freundlichen Grüßen</p><p style="${TEXT_STYLE}">Max Mustermann<br>Tel. +49 30 1234567</p>`,
    );
  });

  it("is plain text: markup is escaped, never taken over", () => {
    const html = signatureHtml({ format: "text", content: '<b>{displayName}</b> <img src=x onerror="a()">' }, profile);
    expect(html).toContain("&lt;b&gt;Max Mustermann&lt;/b&gt; &lt;img src=x onerror=&quot;a()&quot;&gt;");
    expect(html).not.toContain("<b>");
    expect(html).not.toContain("<img");
  });

  it("falls back to MAIL_FOOTER (empty in the tests: no signature)", () => {
    expect(signatureHtml(null, profile)).toBe("");
  });
});
