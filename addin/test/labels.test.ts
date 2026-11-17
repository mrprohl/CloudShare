import { describe, expect, it } from "vitest";
import { parseLabelInfo, parseMsipProperties, readFileLabel } from "../src/taskpane/labels";

// Example IDs only – never the IDs of a real organisation.
const LABEL = "bbbbbbbb-0000-4000-8000-000000000003";
const SITE = "00000000-0000-0000-0000-0000000000aa";
const LABEL_INFO =
  '<?xml version="1.0" encoding="utf-8" standalone="yes"?><clbl:labelList xmlns:clbl="http://schemas.microsoft.com/office/2020/mipLabelMetadata">' +
  `<clbl:label id="{${LABEL.toUpperCase()}}" enabled="1" method="Privileged" siteId="{${SITE}}" contentBits="10" removed="0" /></clbl:labelList>`;

const deflateRaw = async (data: Uint8Array) =>
  new Uint8Array(await new Response(new Blob([data]).stream().pipeThrough(new CompressionStream("deflate-raw"))).arrayBuffer());

/** A ZIP with the given entries (deflated) – enough of an Office file for the label. */
async function zip(files: Record<string, string>): Promise<Blob> {
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const [fileName, text] of Object.entries(files)) {
    const name = new TextEncoder().encode(fileName);
    const raw = new TextEncoder().encode(text);
    const data = await deflateRaw(raw);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(8, 8, true);
    local.setUint32(18, data.length, true);
    local.setUint32(22, raw.length, true);
    local.setUint16(26, name.length, true);
    const dir = new DataView(new ArrayBuffer(46));
    dir.setUint32(0, 0x02014b50, true);
    dir.setUint16(10, 8, true);
    dir.setUint32(20, data.length, true);
    dir.setUint32(24, raw.length, true);
    dir.setUint16(28, name.length, true);
    dir.setUint32(42, offset, true);
    parts.push(new Uint8Array(local.buffer), name, data);
    central.push(new Uint8Array(dir.buffer), name);
    offset += 30 + name.length + data.length;
  }
  const dirSize = central.reduce((sum, p) => sum + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, Object.keys(files).length, true);
  end.setUint16(10, Object.keys(files).length, true);
  end.setUint32(12, dirSize, true);
  end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, new Uint8Array(end.buffer)]);
}

/**
 * An encrypted Office file as Office writes it: a compound file with "EncryptedPackage" (here 5000 bytes, regular
 * sectors) and the label unencrypted in "LabelInfo" (small: in the mini stream).
 */
function encryptedOfficeFile(labelInfo: string | null): Blob {
  const S = 512;
  const END = 0xfffffffe;
  const FREE = 0xffffffff;
  const info = new TextEncoder().encode(labelInfo ?? "");
  const miniSectors = Math.max(1, Math.ceil(info.length / 64));
  const packageSize = 5000;
  const packageSectors = Math.ceil(packageSize / S);
  // Sectors: 0 FAT, 1 directory, 2 mini FAT, 3 mini stream, 4… EncryptedPackage.
  const sectors = 4 + packageSectors;
  const file = new Uint8Array((sectors + 1) * S);
  const v = new DataView(file.buffer);
  [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1].forEach((b, i) => (file[i] = b));
  v.setUint16(0x18, 0x3e, true);
  v.setUint16(0x1a, 3, true);
  v.setUint16(0x1c, 0xfffe, true);
  v.setUint16(0x1e, 9, true);
  v.setUint16(0x20, 6, true);
  v.setUint32(0x2c, 1, true); // FAT sectors
  v.setUint32(0x30, 1, true); // first directory sector
  v.setUint32(0x38, 4096, true); // mini stream cutoff
  v.setUint32(0x3c, 2, true); // first mini FAT sector
  v.setUint32(0x40, 1, true);
  v.setUint32(0x44, END, true); // no DIFAT sectors
  for (let i = 0; i < 109; i++) v.setUint32(0x4c + i * 4, i === 0 ? 0 : FREE, true);
  const sector = (n: number) => (n + 1) * S;
  // FAT
  const fat = Array<number>(S / 4).fill(FREE);
  fat[0] = 0xfffffffd;
  fat[1] = END;
  fat[2] = END;
  fat[3] = END;
  for (let i = 0; i < packageSectors; i++) fat[4 + i] = i === packageSectors - 1 ? END : 5 + i;
  fat.forEach((value, i) => v.setUint32(sector(0) + i * 4, value, true));
  // Mini FAT
  for (let i = 0; i < S / 4; i++) v.setUint32(sector(2) + i * 4, i < miniSectors ? (i === miniSectors - 1 ? END : i + 1) : FREE, true);
  file.set(info, sector(3));
  // Directory
  const entry = (index: number, name: string, type: number, start: number, size: number) => {
    const p = sector(1) + index * 128;
    for (let i = 0; i < name.length; i++) v.setUint16(p + i * 2, name.charCodeAt(i), true);
    v.setUint16(p + 0x40, (name.length + 1) * 2, true);
    v.setUint8(p + 0x42, type);
    v.setUint32(p + 0x44, FREE, true);
    v.setUint32(p + 0x48, FREE, true);
    v.setUint32(p + 0x4c, FREE, true);
    v.setUint32(p + 0x74, start, true);
    v.setUint32(p + 0x78, size, true);
  };
  entry(0, "Root Entry", 5, 3, miniSectors * 64);
  if (labelInfo !== null) entry(1, "LabelInfo", 2, 0, info.length);
  entry(2, "EncryptedPackage", 2, 4, packageSize);
  return new Blob([file]);
}

describe("sensitivity labels – parsing", () => {
  it("reads the label in force from LabelInfo (braces and case do not matter)", () => {
    expect(parseLabelInfo(LABEL_INFO)).toEqual({ id: LABEL, siteId: SITE });
    expect(parseLabelInfo(LABEL_INFO.replace('removed="0"', 'removed="1"'))).toBeNull();
    expect(parseLabelInfo(LABEL_INFO.replace('enabled="1"', 'enabled="0"'))).toBeNull();
  });

  it("reads the MSIP_Label properties and skips a label switched off", () => {
    const on = `<property name="MSIP_Label_${LABEL}_Enabled"><vt:lpwstr>true</vt:lpwstr></property>` +
      `<property name="MSIP_Label_${LABEL}_SiteId"><vt:lpwstr>${SITE}</vt:lpwstr></property>`;
    expect(parseMsipProperties(on)).toEqual({ id: LABEL, siteId: SITE });
    expect(parseMsipProperties(on.replace(">true<", ">false<"))).toBeNull();
    // XMP of a PDF
    expect(parseMsipProperties(`<pdfx:MSIP_Label_${LABEL}_Enabled>true</pdfx:MSIP_Label_${LABEL}_Enabled>`)).toEqual({ id: LABEL, siteId: null });
  });
});

describe("sensitivity labels – files", () => {
  it("Word, Excel, PowerPoint: docMetadata/LabelInfo.xml first, else docProps/custom.xml (deflated)", async () => {
    const current = await zip({ "[Content_Types].xml": "<Types/>", "docMetadata/LabelInfo.xml": LABEL_INFO });
    expect(await readFileLabel(current, "Angebot.docx")).toEqual({ id: LABEL, siteId: SITE, encrypted: false });
    const older = await zip({ "docProps/custom.xml": `<Properties><property name="MSIP_Label_${LABEL}_Enabled"><vt:lpwstr>true</vt:lpwstr></property></Properties>` });
    expect(await readFileLabel(older, "Kalkulation.xlsx")).toEqual({ id: LABEL, siteId: null, encrypted: false });
    expect(await readFileLabel(await zip({ "word/document.xml": "<w:document/>" }), "Ohne.docx")).toBeNull();
  });

  it("encrypted Office file: the label from the unencrypted LabelInfo stream, and encrypted", async () => {
    expect(await readFileLabel(encryptedOfficeFile(LABEL_INFO), "Vertrag.docx")).toEqual({ id: LABEL, siteId: SITE, encrypted: true });
    // Without LabelInfo (older clients): still known as protected.
    expect(await readFileLabel(encryptedOfficeFile(null), "Vertrag.docx")).toEqual({ id: null, siteId: null, encrypted: true });
  });

  it("PDF: the label from the XMP metadata", async () => {
    const pdf = new Blob([`%PDF-1.7\n<x:xmpmeta><pdfx:MSIP_Label_${LABEL}_Enabled>true</pdfx:MSIP_Label_${LABEL}_Enabled>` +
      `<pdfx:MSIP_Label_${LABEL}_SiteId>${SITE}</pdfx:MSIP_Label_${LABEL}_SiteId></x:xmpmeta>\n%%EOF`]);
    expect(await readFileLabel(pdf, "Bericht.pdf")).toEqual({ id: LABEL, siteId: SITE, encrypted: false });
  });

  it("protected files of other types by their extension; plain files and garbage have none", async () => {
    expect(await readFileLabel(new Blob(["x"]), "Plan.pfile")).toEqual({ id: null, siteId: null, encrypted: true });
    expect(await readFileLabel(new Blob(["just text"]), "notes.txt")).toBeNull();
    expect(await readFileLabel(new Blob([new Uint8Array([0x50, 0x4b, 1, 2, 3])]), "broken.zip")).toBeNull();
    expect(await readFileLabel(new Blob([new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0, 0])]), "broken.doc")).toBeNull();
  });
});
