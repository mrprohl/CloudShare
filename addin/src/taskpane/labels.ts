/**
 * Sensitivity labels (Microsoft Purview Information Protection) of the files the user shares – read in the add-in
 * from the file itself, nothing is sent anywhere for it:
 *
 *   Office Open XML (docx, xlsx, pptx, vsdx …) – docMetadata/LabelInfo.xml (current Office) or the custom properties
 *     MSIP_Label_<id>_Enabled / _SiteId in docProps/custom.xml (older Office, other tools);
 *   encrypted Office files and old binary ones (doc, xls, ppt) – an OLE compound file (CFB): an encrypted file keeps
 *     its label unencrypted in the stream "LabelInfo" next to "EncryptedPackage"; old binary files have the
 *     MSIP_Label_… properties in "\x05DocumentSummaryInformation";
 *   PDF – the MSIP_Label_… entries of the XMP metadata;
 *   other files protected by Purview (.pfile, .ppdf, .ptxt, .pjpg …) – recognised as protected, label unknown.
 *
 * Only small parts are read (ZIP directory and two entries, CFB tables and one stream, the ends of a PDF), so even
 * very large files cost little. Anything unreadable simply counts as "no label".
 */

export interface FileLabel {
  /** Label GUID (lower case, no braces); null when only "protected" is known. */
  id: string | null;
  /** Microsoft 365 tenant of the label; null if the file does not say. */
  siteId: string | null;
  /** Encrypted by the label (rights management) – recipients outside the organisation can usually not open it. */
  encrypted: boolean;
}

const GUID = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}";
const cleanGuid = (value: string | undefined | null) => (value ? value.replace(/[{}]/g, "").toLowerCase() : null);

const read = async (blob: Blob, start: number, end: number) =>
  new Uint8Array(await blob.slice(Math.max(0, start), Math.min(blob.size, end)).arrayBuffer());
const latin1 = (bytes: Uint8Array) => {
  let out = "";
  for (let i = 0; i < bytes.length; i += 0x8000) out += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return out;
};
const utf8 = (bytes: Uint8Array) => new TextDecoder("utf-8").decode(bytes);
const utf16 = (bytes: Uint8Array) => new TextDecoder("utf-16le").decode(bytes);

/** `<clbl:label id="{…}" enabled="1" siteId="{…}" removed="0" />` (LabelInfo) – the first label in force. */
export function parseLabelInfo(xml: string): Pick<FileLabel, "id" | "siteId"> | null {
  for (const match of xml.matchAll(/<(?:\w+:)?label\s([^>]*?)\/?>/g)) {
    const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(match[1])?.[1];
    if (attr("enabled") === "0" || attr("removed") === "1") continue;
    const id = cleanGuid(attr("id"));
    if (id) return { id, siteId: cleanGuid(attr("siteId")) };
  }
  return null;
}

/**
 * The MSIP_Label_<id>_Enabled / _SiteId properties (custom.xml, old binary files, PDF metadata) – the first label
 * that is not switched off. Works on any text that contains the names followed by their values.
 */
export function parseMsipProperties(text: string): Pick<FileLabel, "id" | "siteId"> | null {
  for (const match of text.matchAll(new RegExp(`MSIP_Label_(${GUID})_Enabled`, "g"))) {
    // The value follows within a few characters: "…_Enabled"><vt:lpwstr>true<", "…_Enabled>true<", "…_Enabled\0…true".
    const after = text.slice(match.index! + match[0].length, match.index! + match[0].length + 120);
    if (/^[^A-Za-z]*false/i.test(after.replace(/<[^>]*>/g, ""))) continue;
    const id = match[1].toLowerCase();
    const site = new RegExp(`MSIP_Label_${id}_SiteId[^0-9a-fA-F]{0,60}(${GUID})`, "i").exec(text)?.[1];
    return { id, siteId: cleanGuid(site) };
  }
  return null;
}

// --- ZIP (Office Open XML) -------------------------------------------------------------------------------

interface ZipEntry { name: string; method: number; compressedSize: number; offset: number }

async function zipEntries(blob: Blob): Promise<ZipEntry[]> {
  // End of central directory: in the last 64 KiB + 22 bytes.
  const tail = await read(blob, blob.size - 0x10000 - 22, blob.size);
  const view = new DataView(tail.buffer, tail.byteOffset, tail.byteLength);
  let eocd = -1;
  for (let i = tail.length - 22; i >= 0; i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return [];
  let count = view.getUint16(eocd + 10, true);
  let size = view.getUint32(eocd + 12, true);
  let offset = view.getUint32(eocd + 16, true);
  if (offset === 0xffffffff || size === 0xffffffff || count === 0xffff) {
    // ZIP64: the locator right before the end record points to the ZIP64 end record.
    const locator = eocd - 20;
    if (locator < 0 || view.getUint32(locator, true) !== 0x07064b50) return [];
    const recordAt = Number(view.getBigUint64(locator + 8, true));
    const record = await read(blob, recordAt, recordAt + 56);
    const r = new DataView(record.buffer, record.byteOffset, record.byteLength);
    if (r.getUint32(0, true) !== 0x06064b50) return [];
    count = Number(r.getBigUint64(32, true));
    size = Number(r.getBigUint64(40, true));
    offset = Number(r.getBigUint64(48, true));
  }
  if (size > 16 * 1024 * 1024) return []; // no Office file has a directory that large
  const dir = await read(blob, offset, offset + size);
  const d = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries: ZipEntry[] = [];
  for (let p = 0, n = 0; n < count && p + 46 <= dir.length; n++) {
    if (d.getUint32(p, true) !== 0x02014b50) break;
    const nameLength = d.getUint16(p + 28, true);
    const extraLength = d.getUint16(p + 30, true);
    const commentLength = d.getUint16(p + 32, true);
    let compressedSize = d.getUint32(p + 20, true);
    const uncompressedSize = d.getUint32(p + 24, true);
    let localOffset = d.getUint32(p + 42, true);
    const name = utf8(dir.subarray(p + 46, p + 46 + nameLength));
    // ZIP64 extra field: the 64-bit values in the order of the fields that are 0xFFFFFFFF.
    for (let e = p + 46 + nameLength; e + 4 <= p + 46 + nameLength + extraLength; ) {
      const id = d.getUint16(e, true);
      const length = d.getUint16(e + 2, true);
      if (id === 0x0001) {
        let q = e + 4;
        if (uncompressedSize === 0xffffffff) q += 8;
        if (compressedSize === 0xffffffff) {
          compressedSize = Number(d.getBigUint64(q, true));
          q += 8;
        }
        if (localOffset === 0xffffffff) localOffset = Number(d.getBigUint64(q, true));
      }
      e += 4 + length;
    }
    entries.push({ name, method: d.getUint16(p + 10, true), compressedSize, offset: localOffset });
    p += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

async function zipText(blob: Blob, entry: ZipEntry): Promise<string | null> {
  if (entry.compressedSize > 4 * 1024 * 1024) return null;
  const header = await read(blob, entry.offset, entry.offset + 30);
  const h = new DataView(header.buffer, header.byteOffset, header.byteLength);
  if (h.getUint32(0, true) !== 0x04034b50) return null;
  const start = entry.offset + 30 + h.getUint16(26, true) + h.getUint16(28, true);
  const data = await read(blob, start, start + entry.compressedSize);
  if (entry.method === 0) return utf8(data);
  if (entry.method !== 8) return null;
  const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return utf8(new Uint8Array(await new Response(stream).arrayBuffer()));
}

async function ooxmlLabel(blob: Blob): Promise<FileLabel | null> {
  const entries = await zipEntries(blob);
  const find = (name: string) => entries.find((e) => e.name.toLowerCase() === name.toLowerCase());
  const labelInfo = find("docMetadata/LabelInfo.xml");
  if (labelInfo) {
    const xml = await zipText(blob, labelInfo);
    const label = xml ? parseLabelInfo(xml) : null;
    if (label) return { ...label, encrypted: false };
  }
  const custom = find("docProps/custom.xml");
  if (custom) {
    const xml = await zipText(blob, custom);
    const label = xml ? parseMsipProperties(xml) : null;
    if (label) return { ...label, encrypted: false };
  }
  return null;
}

// --- Compound file (encrypted Office files, doc/xls/ppt) -------------------------------------------------

const END_OF_CHAIN = 0xfffffffe;
const FREE = 0xffffffff;

async function compoundFileLabel(blob: Blob): Promise<FileLabel | null> {
  const header = await read(blob, 0, 512);
  const h = new DataView(header.buffer, header.byteOffset, header.byteLength);
  const sectorSize = 1 << h.getUint16(0x1e, true);
  const miniSectorSize = 1 << h.getUint16(0x20, true);
  const fatSectors = h.getUint32(0x2c, true);
  const firstDirSector = h.getUint32(0x30, true);
  const miniCutoff = h.getUint32(0x38, true);
  const firstMiniFat = h.getUint32(0x3c, true);
  let difatSector = h.getUint32(0x44, true);
  if (sectorSize !== 512 && sectorSize !== 4096) return null;
  if (fatSectors > 64 * 1024) return null; // a FAT of more than 256 MB: not a real Office file

  const sectorAt = (n: number) => (n + 1) * sectorSize;
  // Sectors of the FAT: 109 in the header, the rest in a chain of DIFAT sectors.
  const fatList: number[] = [];
  for (let i = 0; i < 109 && fatList.length < fatSectors; i++) fatList.push(h.getUint32(0x4c + i * 4, true));
  for (let guard = 0; fatList.length < fatSectors && difatSector !== END_OF_CHAIN && difatSector !== FREE && guard < 4096; guard++) {
    const s = await read(blob, sectorAt(difatSector), sectorAt(difatSector) + sectorSize);
    const v = new DataView(s.buffer, s.byteOffset, s.byteLength);
    const per = sectorSize / 4 - 1;
    for (let i = 0; i < per && fatList.length < fatSectors; i++) fatList.push(v.getUint32(i * 4, true));
    difatSector = v.getUint32(per * 4, true);
  }
  const fat = new Uint32Array(fatList.length * (sectorSize / 4));
  for (let i = 0; i < fatList.length; i++) {
    const s = await read(blob, sectorAt(fatList[i]), sectorAt(fatList[i]) + sectorSize);
    fat.set(new Uint32Array(s.buffer.slice(s.byteOffset, s.byteOffset + sectorSize)), i * (sectorSize / 4));
  }
  const chain = (start: number, limit = 1 << 20) => {
    const out: number[] = [];
    for (let s = start; s !== END_OF_CHAIN && s < fat.length && out.length < limit; s = fat[s]) out.push(s);
    return out;
  };
  const readChain = async (start: number, size: number) => {
    const parts = await Promise.all(chain(start, Math.ceil(size / sectorSize) + 1).map((s) => read(blob, sectorAt(s), sectorAt(s) + sectorSize)));
    const all = new Uint8Array(parts.length * sectorSize);
    parts.forEach((p, i) => all.set(p, i * sectorSize));
    return all.subarray(0, size);
  };

  // Directory: 128-byte entries – name (UTF-16), type, start sector, size.
  const dirSectors = chain(firstDirSector, 4096);
  const dir = await readChain(firstDirSector, dirSectors.length * sectorSize);
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries: Array<{ name: string; type: number; start: number; size: number }> = [];
  for (let p = 0; p + 128 <= dir.length; p += 128) {
    const nameLength = dv.getUint16(p + 0x40, true);
    const type = dv.getUint8(p + 0x42);
    if (!type || nameLength < 2) continue;
    const name = utf16(dir.subarray(p, p + Math.min(64, nameLength) - 2)).replace(/^[\x00-\x1f]/, "");
    entries.push({ name, type, start: dv.getUint32(p + 0x74, true), size: dv.getUint32(p + 0x78, true) });
  }
  const root = entries.find((e) => e.type === 5);
  const stream = async (name: string) => {
    const entry = entries.find((e) => e.type === 2 && e.name === name);
    if (!entry || entry.size > 4 * 1024 * 1024) return null;
    if (entry.size >= miniCutoff || !root) return readChain(entry.start, entry.size);
    // A small stream lives in the mini stream (the root entry's data), in mini sectors chained by the mini FAT.
    const miniStream = await readChain(root.start, root.size);
    const miniFatSectors = chain(firstMiniFat, 4096);
    const miniFatBytes = await readChain(firstMiniFat, miniFatSectors.length * sectorSize);
    const miniFat = new Uint32Array(miniFatBytes.buffer.slice(miniFatBytes.byteOffset, miniFatBytes.byteOffset + miniFatBytes.byteLength));
    const out = new Uint8Array(entry.size);
    let written = 0;
    for (let s = entry.start; s !== END_OF_CHAIN && s < miniFat.length && written < entry.size; s = miniFat[s]) {
      const piece = miniStream.subarray(s * miniSectorSize, (s + 1) * miniSectorSize);
      out.set(piece.subarray(0, Math.min(piece.length, entry.size - written)), written);
      written += piece.length;
    }
    return out;
  };

  const encrypted = entries.some((e) => e.name === "EncryptedPackage");
  const labelInfo = await stream("LabelInfo");
  if (labelInfo) {
    const label = parseLabelInfo(utf8(labelInfo));
    if (label) return { ...label, encrypted };
  }
  const summary = await stream("DocumentSummaryInformation");
  if (summary) {
    const label = parseMsipProperties(latin1(summary)) ?? parseMsipProperties(utf16(summary));
    if (label) return { ...label, encrypted };
  }
  return encrypted ? { id: null, siteId: null, encrypted: true } : null;
}

// --- PDF -------------------------------------------------------------------------------------------------

async function pdfLabel(blob: Blob): Promise<FileLabel | null> {
  // The XMP metadata is usually near the start or (after incremental saves) near the end.
  const window = 2 * 1024 * 1024;
  const text = latin1(await read(blob, 0, window)) + (blob.size > window ? latin1(await read(blob, blob.size - window, blob.size)) : "");
  const label = parseMsipProperties(text);
  return label ? { ...label, encrypted: false } : null;
}

/** Extensions of files Purview protects as a whole (".pfile" and the "p" variants of common types). */
const PROTECTED_EXTENSIONS = /\.(pfile|ppdf|ptxt|pxml|pjpg|pjpeg|ppng|ptif|ptiff|pbmp|pgif|pjpe|pjfif|pjt)$/i;

/** The label of a file, or null if it has none (or it cannot be read). Never throws. */
export async function readFileLabel(blob: Blob, name: string): Promise<FileLabel | null> {
  try {
    if (PROTECTED_EXTENSIONS.test(name)) return { id: null, siteId: null, encrypted: true };
    const head = await read(blob, 0, 8);
    if (head[0] === 0x50 && head[1] === 0x4b) return await ooxmlLabel(blob);
    if (head[0] === 0xd0 && head[1] === 0xcf && head[2] === 0x11 && head[3] === 0xe0) return await compoundFileLabel(blob);
    if (latin1(head).startsWith("%PDF")) return await pdfLabel(blob);
    return null;
  } catch {
    return null;
  }
}

// --- Labels of the organisation, label of the mail -------------------------------------------------------

export interface LabelDetails {
  id: string;
  name: string;
  /** Position in the label policy: a higher number is a more sensitive label. */
  rank: number;
}

/** The labels published to the user (Outlook's catalog, in the policy's order), or null where Outlook has none. */
export async function labelCatalog(): Promise<LabelDetails[] | null> {
  try {
    if (!Office.context.requirements.isSetSupported("Mailbox", "1.13")) return null;
    const catalog = (Office.context as unknown as { sensitivityLabelsCatalog?: CatalogApi }).sensitivityLabelsCatalog;
    if (!catalog) return null;
    const enabled = await new Promise<boolean>((resolve) =>
      catalog.getIsEnabledAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded && r.value === true)));
    if (!enabled) return null;
    const labels = await new Promise<CatalogLabel[] | null>((resolve) =>
      catalog.getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : null)));
    if (!labels) return null;
    const flat: LabelDetails[] = [];
    const walk = (list: CatalogLabel[], parent?: string) => {
      for (const label of list) {
        flat.push({ id: cleanGuid(label.id)!, name: parent ? `${parent} / ${label.name}` : label.name, rank: flat.length });
        if (label.children?.length) walk(label.children, label.name);
      }
    };
    walk(labels);
    return flat;
  } catch {
    return null;
  }
}

interface CatalogLabel { id: string; name: string; children?: CatalogLabel[] }
interface CatalogApi {
  getIsEnabledAsync(cb: (r: Office.AsyncResult<boolean>) => void): void;
  getAsync(cb: (r: Office.AsyncResult<CatalogLabel[]>) => void): void;
}
interface SensitivityLabelApi {
  getAsync(cb: (r: Office.AsyncResult<string>) => void): void;
  setAsync(label: string, cb: (r: Office.AsyncResult<void>) => void): void;
}

const labelApi = () =>
  (Office.context.mailbox.item as unknown as { sensitivityLabel?: SensitivityLabelApi } | null)?.sensitivityLabel;

/** The label of the mail being written (GUID), "" without one, null where Outlook cannot tell. */
export async function mailLabel(): Promise<string | null> {
  const api = labelApi();
  if (!api) return null;
  return new Promise((resolve) =>
    api.getAsync((r) => resolve(r.status === Office.AsyncResultStatus.Succeeded ? (cleanGuid(r.value) ?? "") : null)));
}

/** Gives the mail this label; throws with Outlook's message if it refuses. */
export async function setMailLabel(id: string): Promise<void> {
  const api = labelApi();
  if (!api) throw new Error("sensitivityLabel not available");
  await new Promise<void>((resolve, reject) =>
    api.setAsync(id, (r) => (r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message ?? "setAsync failed")))));
}
