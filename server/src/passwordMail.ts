/**
 * HTML of the password mail (seen by recipients, so German). Same look as the link card the taskpane inserts
 * (addin/src/taskpane/mailCard.ts – keep colours and spacing in sync): white card, brand accent on the left, file
 * type badges, one button. Tables and inline styles only, so it renders in classic Outlook, Outlook on the web,
 * Gmail and on phones.
 */

const FONT = "'Segoe UI',-apple-system,Helvetica,Arial,sans-serif";
const BRAND = "#0f6cbd";

export const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Colours of the file type badge, close to the Office/Fluent file type colours (as in the link card). */
const TYPE_COLORS: Record<string, string> = {
  pdf: "#c4314b",
  doc: "#185abd", docx: "#185abd", rtf: "#185abd", odt: "#185abd",
  xls: "#107c41", xlsx: "#107c41", csv: "#107c41", ods: "#107c41",
  ppt: "#c43e1c", pptx: "#c43e1c", odp: "#c43e1c",
  jpg: "#8764b8", jpeg: "#8764b8", png: "#8764b8", gif: "#8764b8", heic: "#8764b8", svg: "#8764b8",
  mp4: "#8764b8", mov: "#8764b8", mp3: "#8764b8", wav: "#8764b8",
  zip: "#616161", rar: "#616161", "7z": "#616161", tar: "#616161", gz: "#616161",
};

/** Upload links have no file extension – they get an arrow badge instead of "DATEI". */
function typeBadge(name: string): string {
  const ext = /\.([A-Za-z0-9]{1,8})$/.exec(name)?.[1]?.toLowerCase() ?? "";
  const label = ext ? escapeHtml(ext.toUpperCase().slice(0, 4)) : "&#8593;";
  const color = TYPE_COLORS[ext] ?? (ext ? "#616161" : BRAND);
  return (
    `<td width="32" valign="middle" style="width:32px;padding:4px 10px 4px 0;">` +
    `<div style="width:32px;height:32px;line-height:32px;background:${color};border-radius:4px;color:#ffffff;` +
    `font-family:${FONT};font-size:10px;font-weight:700;text-align:center;">${label}</div></td>`
  );
}

export const accessText = (count?: number) =>
  count === undefined ? "begrenzte Anzahl Abrufe" : count === 1 ? "nur einmal abrufbar" : `bis zu ${count} Abrufe`;

export interface PasswordMailContent {
  fileNames: string[];
  sendUrl: string;
  maxAccessCount?: number;
  /** Subject of the mail with the links ("" if unknown). */
  subject: string;
  /** Days until the password link expires. */
  expireDays: number;
}

export function passwordMailHtml(mail: PasswordMailContent): string {
  const many = mail.fileNames.length !== 1;
  const rows = mail.fileNames
    .map(
      (name) =>
        `<tr>${typeBadge(name)}<td valign="middle" style="padding:4px 0;font-family:${FONT};font-size:14px;` +
        `line-height:20px;font-weight:600;color:#242424;word-break:break-all;">${escapeHtml(name)}</td></tr>`,
    )
    .join("");
  const intro =
    `für ${many ? "die folgenden Freigaben" : "die folgende Freigabe"}` +
    (mail.subject ? ` aus der Nachricht „${escapeHtml(mail.subject)}“` : "") +
    ` benötigen Sie ein Kennwort. Sie können es über den Button abrufen:`;
  const p = (text: string) =>
    `<p style="margin:0 0 12px;font-family:${FONT};font-size:14px;line-height:20px;color:#242424;">${text}</p>`;
  return (
    p("Guten Tag,") +
    p(intro) +
    `<table role="presentation" width="560" cellpadding="0" cellspacing="0" border="0" ` +
    `style="width:100%;max-width:560px;border-collapse:separate;margin:4px 0 12px;">` +
    `<tr><td style="background:#ffffff;border:1px solid #e0e0e0;border-left:4px solid ${BRAND};border-radius:6px;` +
    `padding:14px 16px;font-family:${FONT};color:#242424;">` +
    `<div style="font-size:15px;line-height:20px;font-weight:600;color:#242424;margin:0 0 8px;">` +
    `Kennwort für ${many ? `${mail.fileNames.length} Freigaben` : "Ihre Freigabe"}</div>` +
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">${rows}</table>` +
    `<div style="margin:14px 0 0;">` +
    `<a href="${escapeHtml(mail.sendUrl)}" style="display:inline-block;background:${BRAND};border:1px solid ${BRAND};` +
    `border-radius:4px;color:#ffffff;font-family:${FONT};font-size:13px;font-weight:600;line-height:18px;` +
    `padding:6px 14px;text-decoration:none;white-space:nowrap;">Kennwort anzeigen</a></div>` +
    `<div style="font-size:12px;line-height:16px;color:#616161;margin:10px 0 0;">` +
    `${accessText(mail.maxAccessCount)} · läuft nach ${mail.expireDays} Tagen ab · ` +
    `Geben Sie das Kennwort beim Öffnen des Download-Links ein.</div>` +
    `</td></tr></table>` +
    // The plain link as well – some mail clients block buttons or show them without colours.
    `<p style="margin:0;font-family:${FONT};font-size:12px;line-height:16px;color:#8a8a8a;">` +
    `Falls der Button nicht funktioniert: <a href="${escapeHtml(mail.sendUrl)}" style="color:${BRAND};">` +
    `${escapeHtml(mail.sendUrl)}</a></p>`
  );
}
