/**
 * Signs the delivery proof PDF (PAdES-style, SubFilter adbe.pkcs7.detached) so that any change afterwards shows in
 * every PDF reader ("the document was changed after it was signed"):
 *
 *   identity  – PROOF_SIGN_CERT_FILE + PROOF_SIGN_KEY_FILE (PEM, e.g. a certificate of a CA trusted by Adobe), else a
 *               certificate CloudShare makes itself once (RSA 2048, ten years, CN = PROOF_SIGN_NAME), kept sealed in
 *               the database (store "signing-identity");
 *   signature – a detached CMS SignedData over the PDF bytes outside /Contents, with the signer certificate and the
 *               ESS signing-certificate-v2 attribute;
 *   time      – an RFC 3161 time stamp of the signature from PROOF_TSA_URL (only its hash leaves the server) as unsigned
 *               attribute; "off" or a failing time stamp authority: signed without.
 *
 * pdf.ts leaves a signature field with placeholders (/ByteRange, /Contents); this fills them in place, so the
 * cross-reference table stays right. Built on node:crypto and der.ts only.
 */
import { createHash, createPrivateKey, generateKeyPairSync, randomBytes, sign, X509Certificate, type KeyObject } from "node:crypto";
import { readFileSync } from "node:fs";
import { config } from "./config.js";
import { PersistentMap } from "./db.js";
import { BYTE_RANGE_PLACEHOLDER, SIGNATURE_HEX_LENGTH } from "./pdf.js";
import { bits, bool, children, explicit, implicit, int, nul, octets, oid, read, seq, set, time, utf8 } from "./der.js";

const OID = {
  sha256: "2.16.840.1.101.3.4.2.1",
  rsaEncryption: "1.2.840.113549.1.1.1",
  sha256WithRSA: "1.2.840.113549.1.1.11",
  data: "1.2.840.113549.1.7.1",
  signedData: "1.2.840.113549.1.7.2",
  contentType: "1.2.840.113549.1.9.3",
  messageDigest: "1.2.840.113549.1.9.4",
  signingCertificateV2: "1.2.840.113549.1.9.16.2.47",
  timeStampToken: "1.2.840.113549.1.9.16.2.14",
  commonName: "2.5.4.3",
  organization: "2.5.4.10",
  basicConstraints: "2.5.29.19",
  keyUsage: "2.5.29.15",
};

const sha256 = (data: Buffer) => createHash("sha256").update(data).digest();
const algorithm = (id: string, params = true) => (params ? seq(oid(id), nul()) : seq(oid(id)));

interface Identity {
  key: KeyObject;
  /** The certificate in DER. */
  cert: Buffer;
}

const store = new PersistentMap<{ keyPem: string; certDer: string }>("signing-identity", { retention: false });

/** A self-signed certificate for signing documents: CN (and O), ten years, digital signature + non-repudiation. */
export function makeCertificate(name: string, organization: string, key: KeyObject, publicDer: Buffer, now = new Date()): Buffer {
  const dn = seq(
    set(seq(oid(OID.commonName), utf8(name))),
    ...(organization ? [set(seq(oid(OID.organization), utf8(organization)))] : []),
  );
  const notAfter = new Date(now);
  notAfter.setUTCFullYear(now.getUTCFullYear() + 10);
  const extensions = seq(
    seq(oid(OID.basicConstraints), bool(true), octets(seq())),
    // digitalSignature (bit 0) and nonRepudiation (bit 1): 0b11000000, 6 unused bits.
    seq(oid(OID.keyUsage), bool(true), octets(bits(Buffer.from([0xc0]), 6))),
  );
  const tbs = seq(
    explicit(0, int(2)),
    int(Buffer.concat([Buffer.from([0x01]), randomBytes(15)])),
    algorithm(OID.sha256WithRSA),
    dn,
    seq(time(now), time(notAfter)),
    dn,
    publicDer,
    explicit(3, extensions),
  );
  return seq(tbs, algorithm(OID.sha256WithRSA), bits(sign("sha256", tbs, key)));
}

let cached: Identity | undefined;

/** The signing identity: the configured PEM files, else the one CloudShare made itself (made on first use). */
export function signingIdentity(): Identity {
  if (cached) return cached;
  if (config.proofSignCertFile && config.proofSignKeyFile) {
    const cert = new X509Certificate(readFileSync(config.proofSignCertFile));
    cached = { key: createPrivateKey(readFileSync(config.proofSignKeyFile)), cert: cert.raw };
    return cached;
  }
  const stored = store.get("default");
  if (stored) {
    cached = { key: createPrivateKey(stored.keyPem), cert: Buffer.from(stored.certDer, "base64") };
    return cached;
  }
  const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
  const cert = makeCertificate(config.proofSignName, config.proofSignOrganization, privateKey, publicKey.export({ type: "spki", format: "der" }));
  store.set("default", { keyPem: privateKey.export({ type: "pkcs8", format: "pem" }).toString(), certDer: cert.toString("base64") });
  cached = { key: privateKey, cert };
  return cached;
}

/** Tests only: forget the identity in memory. */
export const resetSigningIdentity = () => {
  cached = undefined;
};

/** RFC 3161: a time stamp token for `signature` (the hash only), or null without a time stamp authority. */
export async function timeStampToken(signature: Buffer): Promise<Buffer | null> {
  const url = config.proofTsaUrl;
  if (!url) return null;
  const request = seq(int(1), seq(algorithm(OID.sha256), octets(sha256(signature))), int(randomBytes(8)), bool(true));
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/timestamp-query", Accept: "application/timestamp-reply" },
    body: request,
    signal: AbortSignal.timeout(15_000),
  });
  if (!res.ok) throw new Error(`time stamp authority: HTTP ${res.status}`);
  const reply = children(read(Buffer.from(await res.arrayBuffer())));
  const status = children(reply[0])[0];
  // 0 granted, 1 granted with modifications – anything else is a refusal.
  if (status.content.length !== 1 || status.content[0] > 1 || !reply[1]) throw new Error("time stamp authority refused the request");
  return Buffer.from(reply[1].raw);
}

/** Issuer and serial number of a certificate (the signer identifier of the CMS). */
function issuerAndSerial(cert: Buffer): Buffer {
  const tbs = children(children(read(cert))[0]);
  const offset = tbs[0].tag === 0xa0 ? 1 : 0; // [0] version
  return seq(tbs[offset].raw, tbs[offset + 2].raw); // issuer is after the signature algorithm
}

/** A detached CMS SignedData over `content`, with a time stamp if one can be had. */
export async function cmsSignature(content: Buffer, identity = signingIdentity()): Promise<Buffer> {
  const signedAttrs = set(
    seq(oid(OID.contentType), set(oid(OID.data))),
    seq(oid(OID.messageDigest), set(octets(sha256(content)))),
    // ESS signing-certificate-v2: binds the signature to this certificate (SHA-256 is the default hash, so it is left out).
    seq(oid(OID.signingCertificateV2), set(seq(seq(seq(octets(sha256(identity.cert))))))),
  );
  const signature = sign("sha256", signedAttrs, identity.key);
  const token = await timeStampToken(signature).catch(() => null);
  const signerInfo = seq(
    int(1),
    issuerAndSerial(identity.cert),
    algorithm(OID.sha256),
    implicit(0, signedAttrs),
    algorithm(OID.rsaEncryption),
    octets(signature),
    ...(token ? [implicit(1, set(seq(oid(OID.timeStampToken), set(token))))] : []),
  );
  const signedData = seq(int(1), set(algorithm(OID.sha256)), seq(oid(OID.data)), implicit(0, set(identity.cert)), set(signerInfo));
  return seq(oid(OID.signedData), explicit(0, signedData));
}

// Placeholders written by pdf.ts into the signature dictionary.

/** Fills the signature of a PDF made with a signature field (pdf.ts `signed`): byte range and CMS. */
export async function signPdf(pdf: Buffer): Promise<Buffer> {
  const text = pdf.toString("latin1");
  const rangeAt = text.indexOf(BYTE_RANGE_PLACEHOLDER);
  const contentsAt = text.indexOf("/Contents <", rangeAt);
  if (rangeAt < 0 || contentsAt < 0) return pdf; // no signature field
  const start = contentsAt + "/Contents ".length; // the "<"
  const end = start + SIGNATURE_HEX_LENGTH + 2; // after the ">"
  // Same length as the placeholder (spaces before "]"), so no offset after it moves.
  const inner = `0 ${start} ${end} ${pdf.length - end}`.padEnd(BYTE_RANGE_PLACEHOLDER.length - "/ByteRange []".length, " ");
  const out = Buffer.from(pdf);
  out.write(`/ByteRange [${inner}]`, rangeAt, "latin1");
  const signed = Buffer.concat([out.subarray(0, start), out.subarray(end)]);
  const hex = (await cmsSignature(signed)).toString("hex");
  if (hex.length > SIGNATURE_HEX_LENGTH) throw new Error("signature larger than its placeholder");
  out.write(hex.padEnd(SIGNATURE_HEX_LENGTH, "0"), start + 1, "latin1");
  return out;
}

/** Tests only: the bytes the signature covers and the CMS in /Contents of a signed PDF. */
export function signedParts(pdf: Buffer): { content: Buffer; cms: Buffer; range: number[] } {
  const text = pdf.toString("latin1");
  const range = /\/ByteRange \[(\d+) (\d+) (\d+) (\d+) *\]/.exec(text)!.slice(1).map(Number);
  const content = Buffer.concat([pdf.subarray(range[0], range[0] + range[1]), pdf.subarray(range[2], range[2] + range[3])]);
  const cms = read(Buffer.from(pdf.subarray(range[1] + 1, range[2] - 1).toString("latin1"), "hex")).raw;
  return { content, cms: Buffer.from(cms), range };
}

