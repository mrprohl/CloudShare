/**
 * A small DER (ASN.1) encoder and reader – enough for the signature of the delivery proof PDF (pdfSign.ts): a self-made
 * X.509 certificate, a CMS SignedData and an RFC 3161 time stamp request and response. No dependency.
 */

const length = (n: number): Buffer => {
  if (n < 0x80) return Buffer.from([n]);
  const bytes: number[] = [];
  for (let v = n; v > 0; v = Math.floor(v / 256)) bytes.unshift(v & 0xff);
  return Buffer.from([0x80 | bytes.length, ...bytes]);
};

/** One TLV: tag byte and content. */
export const tlv = (tag: number, content: Buffer): Buffer => Buffer.concat([Buffer.from([tag]), length(content.length), content]);

export const seq = (...items: Buffer[]) => tlv(0x30, Buffer.concat(items));
export const set = (...items: Buffer[]) => tlv(0x31, Buffer.concat(items));
export const nul = () => Buffer.from([0x05, 0x00]);
export const bool = (value: boolean) => tlv(0x01, Buffer.from([value ? 0xff : 0x00]));
export const octets = (data: Buffer) => tlv(0x04, data);
export const utf8 = (text: string) => tlv(0x0c, Buffer.from(text, "utf8"));
/** BIT STRING with no unused bits (keys, signatures). */
export const bits = (data: Buffer, unused = 0) => tlv(0x03, Buffer.concat([Buffer.from([unused]), data]));
/** [n] EXPLICIT: a constructed context tag around a whole TLV. */
export const explicit = (n: number, inner: Buffer) => tlv(0xa0 | n, inner);
/** [n] IMPLICIT for a constructed value (SET, SEQUENCE): the content of `inner` under the context tag. */
export const implicit = (n: number, inner: Buffer) => tlv(0xa0 | n, read(inner).content);

/** A non-negative INTEGER (big-endian bytes); a leading zero keeps it positive. */
export function int(value: number | bigint | Buffer): Buffer {
  let bytes: Buffer;
  if (Buffer.isBuffer(value)) bytes = value;
  else {
    let hex = BigInt(value).toString(16);
    if (hex.length % 2) hex = `0${hex}`;
    bytes = Buffer.from(hex, "hex");
  }
  let start = 0;
  while (start < bytes.length - 1 && bytes[start] === 0 && bytes[start + 1] < 0x80) start++;
  bytes = bytes.subarray(start);
  if (bytes[0] & 0x80) bytes = Buffer.concat([Buffer.from([0]), bytes]);
  return tlv(0x02, bytes);
}

export function oid(dotted: string): Buffer {
  const parts = dotted.split(".").map(Number);
  const bytes = [40 * parts[0] + parts[1]];
  for (const part of parts.slice(2)) {
    const chunk: number[] = [];
    let v = part;
    do {
      chunk.unshift(v & 0x7f);
      v = Math.floor(v / 128);
    } while (v > 0);
    for (let i = 0; i < chunk.length - 1; i++) chunk[i] |= 0x80;
    bytes.push(...chunk);
  }
  return tlv(0x06, Buffer.from(bytes));
}

/** UTCTime until 2049, GeneralizedTime after (RFC 5280). */
export function time(date: Date): Buffer {
  const iso = date.toISOString().replace(/[-:T]/g, "").slice(0, 14); // YYYYMMDDHHMMSS
  return date.getUTCFullYear() < 2050 ? tlv(0x17, Buffer.from(`${iso.slice(2)}Z`)) : tlv(0x18, Buffer.from(`${iso}Z`));
}

export const generalizedTime = (date: Date) =>
  tlv(0x18, Buffer.from(`${date.toISOString().replace(/[-:T]/g, "").slice(0, 14)}Z`));

export interface Node {
  tag: number;
  /** The whole TLV and its content. */
  raw: Buffer;
  content: Buffer;
}

/** The first TLV in `data`. */
export function read(data: Buffer, offset = 0): Node {
  const tag = data[offset];
  let len = data[offset + 1];
  let head = 2;
  if (len & 0x80) {
    const count = len & 0x7f;
    len = 0;
    for (let i = 0; i < count; i++) len = len * 256 + data[offset + 2 + i];
    head = 2 + count;
  }
  const end = offset + head + len;
  if (end > data.length) throw new Error("DER: length beyond the data");
  return { tag, raw: data.subarray(offset, end), content: data.subarray(offset + head, end) };
}

/** The TLVs inside a constructed value. */
export function children(node: Node): Node[] {
  const out: Node[] = [];
  for (let offset = 0; offset < node.content.length; ) {
    const child = read(node.content, offset);
    out.push(child);
    offset += child.raw.length;
  }
  return out;
}
