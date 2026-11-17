import { describe, expect, it } from "vitest";
import { pdfDocument } from "../src/pdf.js";
import {
  addEvent,
  addLink,
  caseById,
  caseForLink,
  clientInfo,
  endLink,
  eventForTokens,
  linkUrlFor,
  newTrackingId,
  openCase,
  recordMailSent,
  verifyCase,
} from "../src/proof.js";
import { proofPdf } from "../src/proofPdf.js";

const link = (name = "Angebot.pdf") => ({ url: "https://cloud.example.test/s/tok", name, kind: "file" as const, expireDate: "2099-12-31" });

describe("delivery proof – case and chain", () => {
  it("names a case by date, time and four characters without look-alikes", () => {
    const id = newTrackingId(new Date("2026-11-17T10:46:00Z"));
    expect(id).toMatch(/^20261117-1146-[2-9A-HJ-NP-Z]{4}$/); // Europe/Berlin: 11:46
  });

  it("chains every event to the one before and notices a change afterwards", () => {
    const id = openCase("oid-anna", "anna@example.test", { host: "Outlook", platform: "PC" });
    addLink(id, "chain-1", link(), "send-1", { size: 1234, sha256: "ab".repeat(32), protection: "password" });
    addEvent(id, "download", { links: ["chain-1"], downloadedAt: "2026-11-17T12:00:00Z" });
    const c = caseById(id)!;
    expect(c.events.map((e) => e.type)).toEqual(["case.opened", "link.created", "download"]);
    expect(c.events[1].prev).toBe(c.events[0].hash);
    expect(verifyCase(c)).toEqual({ intact: true });

    // Someone edits the record afterwards: the chain no longer fits from that entry on.
    const tampered = structuredClone(c);
    tampered.events[1].data.sha256 = "cd".repeat(32);
    expect(verifyCase(tampered)).toEqual({ intact: false, brokenAt: 1 });
    // Removing an entry breaks it as well.
    const shortened = structuredClone(c);
    shortened.events.splice(1, 1);
    expect(verifyCase(shortened).intact).toBe(false);
  });

  it("keeps all links of a mail in one case until the mail was sent", () => {
    const first = caseForLink(true, undefined, "oid-anna", "anna@example.test", { host: "Outlook" }, {});
    expect(first).toBeDefined();
    expect(caseForLink(true, first, "oid-anna", "anna@example.test", {}, {})).toBe(first);
    // Not ticked: no case; someone else's ID: a new case.
    expect(caseForLink(undefined, first, "oid-anna", "anna@example.test", {}, {})).toBeUndefined();
    expect(caseForLink(true, first, "oid-bernd", "bernd@example.test", {}, {})).not.toBe(first);

    addLink(first!, "same-1", link(), undefined, {});
    recordMailSent(["same-1"], { subject: "Angebot", messageId: "<m1@example.test>", from: "anna@example.test", to: ["Erika@Example.test"], cc: [], bcc: ["x@example.test"] });
    // Recorded once even if the job retries; recipients in lower case (they confirm their address against them).
    recordMailSent(["same-1"], { subject: "Angebot", messageId: "<m1@example.test>", from: "anna@example.test", to: ["Erika@Example.test"], cc: [], bcc: [] });
    const c = caseById(first!)!;
    expect(c.events.filter((e) => e.type === "message.sent")).toHaveLength(1);
    expect(c.recipients).toEqual(["erika@example.test", "x@example.test"]);
    // A new mail with the box ticked gets its own case.
    expect(caseForLink(true, first, "oid-anna", "anna@example.test", {}, {})).not.toBe(first);
  });

  it("records events for the links of a case only, and an ended link once", () => {
    const id = openCase("oid-anna", "anna@example.test", undefined);
    addLink(id, "ev-1", link(), undefined, {});
    eventForTokens(["ev-1", "not-in-a-case"], "password.sent", { channel: "mail", to: "erika@example.test" });
    endLink("ev-1", "revoked");
    endLink("ev-1", "expired");
    const types = caseById(id)!.events.map((e) => e.type);
    expect(types).toEqual(["case.opened", "link.created", "password.sent", "link.revoked"]);
    expect(caseById(id)!.events[2].data.links).toEqual(["ev-1"]);
  });

  it("puts the gateway into the mail for Nextcloud links of a case, not for others", () => {
    const id = openCase("oid-anna", "anna@example.test", undefined);
    addLink(id, "gw-1", link(), undefined, {});
    expect(linkUrlFor("gw-1", "https://cloud.example.test/s/gw-1")).toBe("https://cloudshare.example.test/r/s/gw-1");
    expect(linkUrlFor("other", "https://cloud.example.test/s/other")).toBe("https://cloud.example.test/s/other");
  });

  it("keeps only short plain values of the client", () => {
    expect(clientInfo({ host: "Outlook", platform: "PC", version: "16.0.18129.20000", extra: "x" })).toEqual({
      host: "Outlook", platform: "PC", version: "16.0.18129.20000", language: undefined,
    });
    expect(clientInfo({ host: "<script>" }).host).toBeUndefined();
  });
});

describe("delivery proof – PDF", () => {
  it("writes a valid PDF with the standard fonts, umlauts in WinAnsi and a footer on every page", () => {
    const lines = Array.from({ length: 120 }, (_, i) => ({ text: `Zeile ${i} – Übergabe größer als Maß` }));
    const pdf = pdfDocument(lines, { title: "Zustellnachweis", footer: (page, pages) => `Seite ${page} von ${pages}` }).toString("latin1");
    expect(pdf.startsWith("%PDF-1.4")).toBe(true);
    expect(pdf.trimEnd().endsWith("%%EOF")).toBe(true);
    expect(pdf).toContain("/BaseFont /Helvetica");
    expect(pdf).toContain("Übergabe größer"); // Ü, ö, ß as single WinAnsi bytes
    expect(pdf).toContain("\x96"); // the en dash
    const pages = Number(/\/Count (\d+)/.exec(pdf)![1]);
    expect(pages).toBeGreaterThan(1);
    expect(pdf).toContain(`(Seite ${pages} von ${pages})`);
    // The cross-reference table points at the objects.
    const xref = Number(/startxref\n(\d+)/.exec(pdf)![1]);
    expect(pdf.slice(xref, xref + 4)).toBe("xref");
  });

  it("escapes brackets and replaces characters the fonts do not have", () => {
    const pdf = pdfDocument([{ text: "a (b) \\ c 漢" }], { title: "x" }).toString("latin1");
    expect(pdf).toContain("(a \\(b\\) \\\\ c ?)");
  });

  it("lists the case with its events, the mail and whether the record is intact", () => {
    const id = openCase("oid-anna", "anna@example.test", { host: "Outlook", platform: "PC", version: "16.0" });
    addLink(id, "pdf-1", link("Angebot (final).pdf"), "send-9", { size: 2_621_440, sha256: "ef".repeat(32), protection: "password", passwordWays: ["recipients", "sms:+49 170 •••• 67"] });
    recordMailSent(["pdf-1"], { subject: "Unterlagen", messageId: "<abc@example.test>", from: "anna@example.test", to: ["erika@example.test"], cc: [], bcc: [], sentAt: "2026-11-17T10:47:00Z" });
    const pdf = proofPdf(caseById(id)!, "de").toString("latin1");
    expect(pdf).toContain(`(Vorgang ${id})`);
    expect(pdf).toContain("Protokoll unverändert: 3 Einträge");
    expect(pdf).toContain("Angebot \\(final\\).pdf");
    expect(pdf).toContain("ef".repeat(16)); // the fingerprint (wrapped lines keep it in pieces)
    expect(pdf).toContain("<abc@example.test>");
    expect(pdf).toContain("SMS an +49 170 \x95\x95\x95\x95 67");
    expect(proofPdf(caseById(id)!, "en").toString("latin1")).toContain(`(Case ${id})`);
  });
});
