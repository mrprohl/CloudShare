/**
 * A minimal PDF writer for text documents (the "Zustellnachweis" PDF, proofPdf.ts) – no dependency: A4 pages, the
 * standard fonts Helvetica, Helvetica-Bold and Courier in WinAnsi encoding (German and Western European text;
 * other characters become "?"), lines wrapped at word boundaries, a footer on every page.
 */

export interface PdfLine {
  text: string;
  /** Font size in points (default 10). */
  size?: number;
  font?: "regular" | "bold" | "mono";
  /** Grey level of the text: 0 black … 1 white (default 0). */
  grey?: number;
  /** Left indent in points. */
  indent?: number;
  /** Space above the line in points. */
  gapBefore?: number;
  /** A thin rule above the line. */
  rule?: boolean;
}

const PAGE_WIDTH = 595.28;
const PAGE_HEIGHT = 841.89;
const MARGIN = 56;
const FOOTER = 36;

/** Helvetica widths (1/1000 em) for the printable ASCII characters 32…126 (Adobe AFM). */
const HELVETICA = [
  278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556,
  556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778,
  667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556,
  556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584,
];

/** WinAnsi codes of the characters outside Latin-1 that the PDF fonts have. */
const WIN_ANSI: Record<string, number> = {
  "€": 0x80, "‚": 0x82, "„": 0x84, "…": 0x85, "‘": 0x91, "’": 0x92, "“": 0x93, "”": 0x94, "•": 0x95, "–": 0x96, "—": 0x97,
  "™": 0x99,
};

/** The text as WinAnsi bytes (one char per byte); characters the fonts lack become "?". */
function winAnsi(text: string): string {
  let out = "";
  for (const ch of text.normalize("NFC")) {
    const code = ch.codePointAt(0)!;
    if (WIN_ANSI[ch]) out += String.fromCharCode(WIN_ANSI[ch]);
    else if ((code >= 32 && code < 127) || (code >= 0xa0 && code <= 0xff)) out += ch;
    else if (ch === "\t") out += " ";
    else out += "?";
  }
  return out;
}

/** Width of a character in 1/1000 em: accented letters as their base letter, the rest as an average letter. */
function charWidth(byte: string, font: PdfLine["font"]): number {
  if (font === "mono") return 600;
  const code = byte.charCodeAt(0);
  let width: number;
  if (code >= 32 && code < 127) width = HELVETICA[code - 32];
  else if (code === 0xdf) width = 611; // ß
  else {
    const base = byte.normalize("NFD")[0];
    const b = base.charCodeAt(0);
    width = b >= 32 && b < 127 ? HELVETICA[b - 32] : 556;
  }
  return font === "bold" ? width * 1.07 : width;
}

const textWidth = (bytes: string, line: PdfLine) =>
  ([...bytes].reduce((sum, ch) => sum + charWidth(ch, line.font), 0) * (line.size ?? 10)) / 1000;

/** Splits a line into pieces that fit `width`: at spaces, a word longer than the line at any character. */
function wrap(bytes: string, line: PdfLine, width: number): string[] {
  const out: string[] = [];
  let current = "";
  for (const word of bytes.split(" ")) {
    const candidate = current ? `${current} ${word}` : word;
    if (textWidth(candidate, line) <= width) {
      current = candidate;
      continue;
    }
    if (current) out.push(current);
    current = "";
    let rest = word;
    while (textWidth(rest, line) > width) {
      let cut = rest.length;
      while (cut > 1 && textWidth(rest.slice(0, cut), line) > width) cut--;
      out.push(rest.slice(0, cut));
      rest = rest.slice(cut);
    }
    current = rest;
  }
  out.push(current);
  return out;
}

const escapePdf = (bytes: string) => bytes.replace(/[\\()]/g, (c) => `\\${c}`);

/** A text string for the document information (UTF-16BE, so any character works). */
function infoString(text: string): string {
  const units = [0xfeff, ...[...text].flatMap((ch) => {
    const code = ch.codePointAt(0)!;
    if (code < 0x10000) return [code];
    const v = code - 0x10000;
    return [0xd800 + (v >> 10), 0xdc00 + (v & 0x3ff)];
  })];
  return `<${units.map((u) => u.toString(16).padStart(4, "0")).join("")}>`;
}

const FONT_NAMES = { regular: "F1", bold: "F2", mono: "F3" } as const;

/** Builds the PDF: `lines` flow over as many pages as needed; `footer(page, pages)` is printed on every page. */
export function pdfDocument(
  lines: PdfLine[],
  meta: { title: string; author?: string; created?: Date; footer?: (page: number, pages: number) => string },
): Buffer {
  const pages: string[][] = [[]];
  let y = PAGE_HEIGHT - MARGIN;
  const bottom = MARGIN + FOOTER;

  for (const line of lines) {
    const size = line.size ?? 10;
    const leading = size * 1.35;
    const indent = line.indent ?? 0;
    const pieces = wrap(winAnsi(line.text), line, PAGE_WIDTH - 2 * MARGIN - indent);
    y -= line.gapBefore ?? 0;
    if (line.rule) {
      if (y - leading < bottom) {
        pages.push([]);
        y = PAGE_HEIGHT - MARGIN;
      }
      pages.at(-1)!.push(`0.8 G 0.5 w ${MARGIN} ${(y + 2).toFixed(2)} m ${PAGE_WIDTH - MARGIN} ${(y + 2).toFixed(2)} l S`);
      y -= 6;
    }
    for (const piece of pieces) {
      if (y - leading < bottom) {
        pages.push([]);
        y = PAGE_HEIGHT - MARGIN;
      }
      y -= leading;
      pages.at(-1)!.push(
        `BT /${FONT_NAMES[line.font ?? "regular"]} ${size} Tf ${line.grey ?? 0} g ${(MARGIN + indent).toFixed(2)} ${y.toFixed(2)} Td (${escapePdf(piece)}) Tj ET`,
      );
    }
  }

  const total = pages.length;
  if (meta.footer) {
    pages.forEach((ops, i) => {
      const text = escapePdf(winAnsi(meta.footer!(i + 1, total)));
      ops.push(`BT /F1 8 Tf 0.45 g ${MARGIN} ${MARGIN - 8} Td (${text}) Tj ET`);
    });
  }

  // Objects: 1 catalog, 2 page tree, 3–5 fonts, 6 info, then per page its page object and content stream.
  const objects: string[] = [];
  const pageIds = pages.map((_, i) => 7 + i * 2);
  objects[1] = "<< /Type /Catalog /Pages 2 0 R >>";
  objects[2] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${total} >>`;
  const font = (name: string) => `<< /Type /Font /Subtype /Type1 /BaseFont /${name} /Encoding /WinAnsiEncoding >>`;
  objects[3] = font("Helvetica");
  objects[4] = font("Helvetica-Bold");
  objects[5] = font("Courier");
  const created = meta.created ?? new Date();
  const date = `D:${created.toISOString().replace(/[-:T]/g, "").slice(0, 14)}Z`;
  objects[6] = `<< /Title ${infoString(meta.title)}${meta.author ? ` /Author ${infoString(meta.author)}` : ""} /Producer (CloudShare) /CreationDate (${date}) >>`;
  pages.forEach((ops, i) => {
    const content = ops.join("\n");
    objects[pageIds[i]] =
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_WIDTH} ${PAGE_HEIGHT}] ` +
      `/Resources << /Font << /F1 3 0 R /F2 4 0 R /F3 5 0 R >> >> /Contents ${pageIds[i] + 1} 0 R >>`;
    objects[pageIds[i] + 1] = `<< /Length ${content.length} >>\nstream\n${content}\nendstream`;
  });

  let body = "%PDF-1.4\n%\xe2\xe3\xcf\xd3\n";
  const offsets: number[] = [];
  for (let id = 1; id < objects.length; id++) {
    offsets[id] = body.length;
    body += `${id} 0 obj\n${objects[id]}\nendobj\n`;
  }
  const xref = body.length;
  body += `xref\n0 ${objects.length}\n0000000000 65535 f \n`;
  for (let id = 1; id < objects.length; id++) body += `${String(offsets[id]).padStart(10, "0")} 00000 n \n`;
  body += `trailer\n<< /Size ${objects.length} /Root 1 0 R /Info 6 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(body, "latin1");
}
