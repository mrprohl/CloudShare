/**
 * HTML cards inserted into the mail body (seen by recipients, so German).
 * Built for mail clients: tables and inline styles only, no classes, no external images. Classic Outlook (Word
 * engine) ignores rounded corners and some padding – the card then degrades to a plain bordered box.
 */

export interface CardFile {
  name: string;
  size: number;
  url: string;
}

const FONT = "'Segoe UI',-apple-system,Helvetica,Arial,sans-serif";
const BRAND = "#0f6cbd";

const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Colours of the file type badge, close to the Office/Fluent file type colours. */
const TYPE_COLORS: Record<string, string> = {
  pdf: "#c4314b",
  doc: "#185abd", docx: "#185abd", rtf: "#185abd", odt: "#185abd",
  xls: "#107c41", xlsx: "#107c41", csv: "#107c41", ods: "#107c41",
  ppt: "#c43e1c", pptx: "#c43e1c", odp: "#c43e1c",
  jpg: "#8764b8", jpeg: "#8764b8", png: "#8764b8", gif: "#8764b8", heic: "#8764b8", svg: "#8764b8",
  mp4: "#8764b8", mov: "#8764b8", mp3: "#8764b8", wav: "#8764b8",
  zip: "#616161", rar: "#616161", "7z": "#616161", tar: "#616161", gz: "#616161",
};

function typeBadge(name: string): string {
  const ext = name.includes(".") ? name.split(".").pop()!.toLowerCase() : "";
  const label = escapeHtml((ext || "datei").toUpperCase().slice(0, 4));
  const color = TYPE_COLORS[ext] ?? "#616161";
  return (
    `<td width="40" valign="middle" style="width:40px;padding:0 12px 0 0;">` +
    `<div style="width:40px;height:40px;line-height:40px;background:${color};border-radius:4px;color:#ffffff;` +
    `font-family:${FONT};font-size:11px;font-weight:700;text-align:center;">${label}</div></td>`
  );
}

function button(url: string, text: string): string {
  return (
    `<a href="${escapeHtml(url)}" style="display:inline-block;background:${BRAND};border:1px solid ${BRAND};` +
    `border-radius:4px;color:#ffffff;font-family:${FONT};font-size:13px;font-weight:600;line-height:18px;` +
    `padding:6px 14px;text-decoration:none;white-space:nowrap;">${text}</a>`
  );
}

/** Outer frame: white card, thin border, brand accent on the left. */
function frame(title: string, rows: string, note: string, footer: string): string {
  return (
    `<table role="presentation" width="560" cellpadding="0" cellspacing="0" border="0" ` +
    `style="width:100%;max-width:560px;border-collapse:separate;margin:12px 0;">` +
    `<tr><td style="background:#ffffff;border:1px solid #e0e0e0;border-left:4px solid ${BRAND};border-radius:6px;` +
    `padding:14px 16px;font-family:${FONT};color:#242424;">` +
    `<div style="font-size:15px;line-height:20px;font-weight:600;color:#242424;margin:0 0 6px;">${title}</div>` +
    rows +
    `<div style="font-size:12px;line-height:16px;color:#616161;margin:10px 0 0;">${note}</div>` +
    (footer
      ? `<div style="font-size:11px;line-height:15px;color:#8a8a8a;margin:10px 0 0;border-top:1px solid #f0f0f0;` +
        `padding-top:8px;">${escapeHtml(footer).split("\n").join("<br>")}</div>`
      : "") +
    `</td></tr></table>`
  );
}

const PASSWORD_NOTE = "Das Kennwort erhalten Sie in einer separaten Nachricht.";

/** Card for shared files: one row per file with type badge, name, size and a download button. */
export function filesCard(
  files: CardFile[],
  expireDate: string,
  footer: string,
  formatSize: (bytes: number) => string,
): string {
  const rows = files
    .map(
      (f, i) =>
        `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" ` +
        `style="width:100%;border-collapse:collapse;${i ? "border-top:1px solid #f0f0f0;" : ""}">` +
        `<tr>${typeBadge(f.name)}` +
        `<td valign="middle" style="padding:8px 12px 8px 0;font-family:${FONT};">` +
        `<div style="font-size:14px;line-height:20px;font-weight:600;color:#242424;word-break:break-all;">` +
        `${escapeHtml(f.name)}</div>` +
        `<div style="font-size:12px;line-height:16px;color:#616161;">${escapeHtml(formatSize(f.size))}</div></td>` +
        `<td valign="middle" align="right" style="padding:8px 0;">${button(f.url, "Herunterladen")}</td>` +
        `</tr></table>`,
    )
    .join("");
  const title = files.length === 1 ? "Eine Datei für Sie bereitgestellt" : `${files.length} Dateien für Sie bereitgestellt`;
  return frame(title, rows, `Gültig bis ${expireDate} · ${PASSWORD_NOTE}`, footer);
}

/** Card for an upload link ("Dateien anfordern"). */
export function requestCard(url: string, expireDate: string, footer: string): string {
  const rows =
    `<div style="font-size:14px;line-height:20px;color:#424242;margin:0 0 12px;">` +
    `Über den folgenden Link können Sie mir Ihre Dateien sicher übermitteln.</div>` +
    button(url, "Dateien hochladen");
  return frame("Dateien sicher hochladen", rows, `Gültig bis ${expireDate} · ${PASSWORD_NOTE}`, footer);
}

/** Plain-text versions for mails composed in text format. */
const textFooter = (footer: string) => (footer ? `--\n${footer}\n` : "");

export function filesText(
  files: CardFile[],
  expireDate: string,
  footer: string,
  formatSize: (bytes: number) => string,
): string {
  const lines = files.map((f) => `- ${f.name} (${formatSize(f.size)}): ${f.url}`);
  return (
    `\nDateien für Sie bereitgestellt (gültig bis ${expireDate}):\n${lines.join("\n")}\n${PASSWORD_NOTE}\n` +
    textFooter(footer)
  );
}

export function requestText(url: string, expireDate: string, footer: string): string {
  return (
    `\nÜber diesen Link können Sie mir Ihre Dateien sicher übermitteln (gültig bis ${expireDate}):\n${url}\n` +
    `${PASSWORD_NOTE}\n` +
    textFooter(footer)
  );
}
