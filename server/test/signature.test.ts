/**
 * Signature of the password mails (src/signature.ts): profile placeholders as MAIL_FOOTER always had them, and the
 * allow-list for formatted signatures – the server builds the HTML from a node tree and never trusts the add-in.
 */
import { describe, expect, it } from "vitest";
import { cleanSignatureNodes, fillTemplate, renderSignatureNodes, signatureHasContent, signatureHtml } from "../src/signature.js";

const profile = {
  displayName: "Max Mustermann",
  jobTitle: "Projektleitung",
  company: "Muster GmbH",
  email: "max@example.test",
  phone: "+49 30 1234567",
  department: "",
};
const TEMPLATE = "{displayName} · {jobTitle}\\n{company} · Tel. {phone} · {email}";

describe("fillTemplate (MAIL_FOOTER and text signatures)", () => {
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

describe("formatted signatures", () => {
  it("keeps formatting, links and https pictures, drops everything else", () => {
    const nodes = cleanSignatureNodes([
      { tag: "p", attrs: { style: "color:#1f497d;font-weight:bold;background:url(https://x.test/t.gif)", onclick: "a()", class: "Mso" }, children: ["Max"] },
      { tag: "script", children: ["alert(1)"] },
      { tag: "a", attrs: { href: "javascript:alert(2)" }, children: ["klick"] },
      { tag: "a", attrs: { href: "mailto:max@example.test" }, children: ["Mail"] },
      { tag: "img", attrs: { src: "cid:image001.png" } },
      { tag: "img", attrs: { src: "https://example.com/logo.png", alt: "Logo", onerror: "a()" } },
      { tag: "marquee", children: ["text stays"] },
      42,
      null,
    ]);
    const html = renderSignatureNodes(nodes);
    expect(html).toContain('<p style="color:#1f497d;font-weight:bold">Max</p>');
    expect(html).toContain("alert(1)"); // the text of an unknown element stays as text …
    expect(html).not.toContain("<script"); // … the element does not
    expect(html).toContain("<a>klick</a>"); // the unsafe href is gone
    expect(html).toContain('<a href="mailto:max@example.test" target="_blank" rel="noopener">Mail</a>');
    expect(html).toContain('<img src="https://example.com/logo.png" alt="Logo">');
    expect(html).not.toMatch(/onclick|onerror|javascript:|cid:|url\(|class=|marquee/);
    expect(html).toContain("text stays");
  });

  it("escapes texts and attribute values", () => {
    const html = renderSignatureNodes([{ tag: "span", attrs: { style: "color:red" }, children: ['<b>"x"</b> & y'] }]);
    expect(html).toBe('<span style="color:red">&lt;b&gt;&quot;x&quot;&lt;/b&gt; &amp; y</span>');
    expect(renderSignatureNodes([{ tag: "a", attrs: { href: 'https://x.test/"><script>' }, children: ["l"] }])).toContain(
      'href="https://x.test/&quot;&gt;&lt;script&gt;"',
    );
  });

  it("knows whether anything is left", () => {
    expect(signatureHasContent(cleanSignatureNodes([{ tag: "p", children: ["  "] }]))).toBe(false);
    expect(signatureHasContent(cleanSignatureNodes([{ tag: "img", attrs: { src: "https://x.test/l.png" } }]))).toBe(true);
  });
});

describe("signatureHtml", () => {
  it("uses the own formatted or text signature, else MAIL_FOOTER (empty in the tests)", () => {
    expect(signatureHtml({ format: "html", nodes: [{ tag: "b", children: ["Max"] }] }, profile)).toBe("<b>Max</b>");
    expect(signatureHtml({ format: "text", content: "{displayName}\nTel. {phone}" }, profile)).toBe("Max Mustermann<br>Tel. +49 30 1234567");
    expect(signatureHtml({ format: "text", content: "<b>{displayName}</b>" }, profile)).toBe("&lt;b&gt;Max Mustermann&lt;/b&gt;");
    expect(signatureHtml(null, profile)).toBe("");
  });
});
