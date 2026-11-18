import { spawnSync } from "node:child_process";
import { generateKeyPairSync, X509Certificate } from "node:crypto";
import { mkdtempSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { config } from "../src/config.js";
import { children, explicit, int, oid, read, seq } from "../src/der.js";
import { pdfDocument } from "../src/pdf.js";
import { makeCertificate, signedParts, signingIdentity, signPdf } from "../src/pdfSign.js";
import { mockFetch } from "./helpers.js";

afterEach(() => {
  vi.unstubAllGlobals();
  config.proofTsaUrl = "";
});

const pdf = () =>
  pdfDocument([{ text: "Zustellnachweis – Vorgang 20261118-1646-CBMH" }], { title: "Zustellnachweis", signature: { name: "CloudShare", reason: "Test" } });

/** Whether OpenSSL can check the CMS (it is on the CI runners; skipped where it is missing). */
function opensslVerify(cms: Buffer, content: Buffer): boolean | null {
  const dir = mkdtempSync(path.join(tmpdir(), "cloudshare-sig-"));
  writeFileSync(path.join(dir, "sig.der"), cms);
  writeFileSync(path.join(dir, "data.bin"), content);
  const run = spawnSync("openssl", ["cms", "-verify", "-binary", "-inform", "DER", "-in", path.join(dir, "sig.der"),
    "-content", path.join(dir, "data.bin"), "-noverify", "-out", path.join(dir, "out.bin")]);
  if (run.error) return null;
  return run.status === 0;
}

describe("signing the proof PDF", () => {
  it("makes a certificate Node reads and whose signature checks", () => {
    const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
    const der = makeCertificate("CloudShare Zustellnachweis", "Example GmbH", privateKey, publicKey.export({ type: "spki", format: "der" }));
    const cert = new X509Certificate(der);
    expect(cert.subject).toContain("CN=CloudShare Zustellnachweis");
    expect(cert.subject).toContain("O=Example GmbH");
    expect(cert.verify(publicKey)).toBe(true);
    expect(new Date(cert.validTo).getTime()).toBeGreaterThan(Date.now() + 9 * 365 * 86_400_000);
  });

  it("keeps its own identity once made", () => {
    expect(signingIdentity().cert.equals(signingIdentity().cert)).toBe(true);
  });

  it("signs the bytes outside /Contents; a change afterwards breaks the signature", async () => {
    const signed = await signPdf(pdf());
    const text = signed.toString("latin1");
    expect(text).toContain("/SubFilter /adbe.pkcs7.detached");
    expect(text).toContain("/AcroForm << /Fields [");
    // The cross-reference table still points at the objects (same length as before signing).
    const xref = Number(/startxref\n(\d+)/.exec(text)![1]);
    expect(text.slice(xref, xref + 4)).toBe("xref");

    const { content, cms, range } = signedParts(signed);
    expect(range[0]).toBe(0);
    expect(range[2] + range[3]).toBe(signed.length);
    // The CMS is a SignedData with our certificate.
    expect(children(read(cms))[0].raw.equals(oid("1.2.840.113549.1.7.2"))).toBe(true);
    expect(cms.includes(signingIdentity().cert)).toBe(true);

    const verified = opensslVerify(cms, content);
    if (verified !== null) {
      expect(verified).toBe(true);
      const changed = Buffer.from(content);
      changed[changed.indexOf("Vorgang") + 1] ^= 0x01; // one changed letter
      expect(opensslVerify(cms, changed)).toBe(false);
    }
  });

  it("adds the time stamp of the authority as an unsigned attribute (only the hash goes out)", async () => {
    config.proofTsaUrl = "https://tsa.example.test";
    // A granted TimeStampResp: status 0 and a (fake) token ContentInfo.
    const token = seq(oid("1.2.840.113549.1.7.2"), explicit(0, seq(int(3))));
    const { call } = mockFetch(new Response(seq(seq(int(0)), token), { status: 200, headers: { "Content-Type": "application/timestamp-reply" } }));
    const signed = await signPdf(pdf());
    const { cms, content } = signedParts(signed);
    expect(call(0).url).toBe("https://tsa.example.test");
    expect((call(0).init.headers as Record<string, string>)["Content-Type"]).toBe("application/timestamp-query");
    // The request carries a SHA-256 hash (32 bytes), never the PDF.
    expect(Buffer.from(call(0).init.body as Uint8Array).length).toBeLessThan(200);
    expect(cms.includes(oid("1.2.840.113549.1.9.16.2.14"))).toBe(true);
    expect(cms.includes(token)).toBe(true);
    const verified = opensslVerify(cms, content);
    if (verified !== null) expect(verified).toBe(true);
  });

  it("signs without a time stamp when the authority fails", async () => {
    config.proofTsaUrl = "https://tsa.example.test";
    mockFetch(new Response("down", { status: 503 }));
    const { cms } = signedParts(await signPdf(pdf()));
    expect(cms.includes(oid("1.2.840.113549.1.9.16.2.14"))).toBe(false);
  });
});
