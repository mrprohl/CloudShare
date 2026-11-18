/**
 * One layout for every mail CloudShare sends itself – the password mail to the recipients, the download confirmation
 * and the notes about a password SMS to the sender. Same look as the link card the taskpane inserts
 * (addin/src/taskpane/mailCard.ts – keep colours and spacing in sync):
 *
 *   text paragraphs · one white card (accent on the left, title, file rows with type badges, at most one button or
 *   numbered steps each with its button and plain link, a small detail line) · optional text below · grey line "sent automatically by CloudShare"
 *
 * The colour of the accent tells the kind: brand blue (password), green (downloaded), red (problem). Tables and inline
 * styles only, so it renders in classic Outlook, Outlook on the web, Gmail and on phones.
 */
import { t } from "./i18n.js";

export const FONT = "'Segoe UI',-apple-system,Helvetica,Arial,sans-serif";
export const BRAND = "#0f6cbd";
const ACCENTS = { brand: BRAND, success: "#107c10", danger: "#c50f1f" } as const;
export type Accent = keyof typeof ACCENTS;

export const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Style of a text paragraph – also used for the paragraphs of the text templates (templates.ts, renderHtml). */
export const TEXT_STYLE = `margin:0 0 12px;font-family:${FONT};font-size:14px;line-height:20px;color:#242424;`;

/** A paragraph of plain text (escaped here). */
export const paragraph = (text: string) => `<p style="${TEXT_STYLE}">${escapeHtml(text)}</p>`;

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

const badgeCell = (label: string, color: string) =>
  `<td width="32" valign="middle" style="width:32px;padding:4px 10px 4px 0;">` +
  `<div style="width:32px;height:32px;line-height:32px;background:${color};border-radius:4px;color:#ffffff;` +
  `font-family:${FONT};font-size:10px;font-weight:700;text-align:center;">${label}</div></td>`;

/** File type badge; upload links (no extension) get an arrow, confidential messages an envelope. */
function typeBadge(name: string): string {
  if (name === t("server.messageItem")) return badgeCell("&#9993;", BRAND);
  const ext = /\.([A-Za-z0-9]{1,8})$/.exec(name)?.[1]?.toLowerCase() ?? "";
  const label = ext ? escapeHtml(ext.toUpperCase().slice(0, 4)) : "&#8593;";
  return badgeCell(label, TYPE_COLORS[ext] ?? (ext ? "#616161" : BRAND));
}

export interface MailCard {
  accent: Accent;
  title: string;
  /** One row per file (upload link, confidential message) with its type badge. */
  files: string[];
  /** A short value shown large under the title (the six-digit code of the delivery proof). */
  highlight?: string;
  button?: { label: string; url: string };
  /**
   * Numbered steps instead of the one button (the password mail after the delivery proof's confirmation): each with
   * its text, its button and, right below, the same address as plain text for clients that hide buttons.
   */
  steps?: { text: string; button: { label: string; url: string } }[];
  /** Small grey line at the bottom of the card (validity, accesses, number …). */
  details?: string;
}

/** The card: white box, coloured accent on the left, title, file rows, optional button and detail line. */
export function mailCard(card: MailCard): string {
  const rows = card.files
    .map(
      (name) =>
        `<tr>${typeBadge(name)}<td valign="middle" style="padding:4px 0;font-family:${FONT};font-size:14px;` +
        `line-height:20px;font-weight:600;color:#242424;word-break:break-all;">${escapeHtml(name)}</td></tr>`,
    )
    .join("");
  const button = card.button ? `<div style="margin:14px 0 0;">${buttonLink(card.button)}</div>` : "";
  const steps = (card.steps ?? [])
    .map(
      (step, i) =>
        `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:14px 0 0;width:100%;"><tr>` +
        `<td width="24" valign="top" style="width:24px;padding:0 10px 0 0;">` +
        `<div style="width:24px;height:24px;line-height:24px;border-radius:12px;background:${BRAND};color:#ffffff;` +
        `font-family:${FONT};font-size:12px;font-weight:700;text-align:center;">${i + 1}</div></td>` +
        `<td valign="top" style="font-family:${FONT};">` +
        `<div style="font-size:14px;line-height:24px;font-weight:600;color:#242424;margin:0 0 6px;">${escapeHtml(step.text)}</div>` +
        buttonLink(step.button) +
        `<div style="font-size:12px;line-height:16px;margin:6px 0 0;word-break:break-all;">` +
        `<a href="${escapeHtml(step.button.url)}" style="color:${BRAND};">${escapeHtml(step.button.url)}</a></div>` +
        `</td></tr></table>`,
    )
    .join("");
  const details = card.details
    ? `<div style="font-size:12px;line-height:16px;color:#616161;margin:10px 0 0;">${escapeHtml(card.details)}</div>`
    : "";
  return (
    `<table role="presentation" width="560" cellpadding="0" cellspacing="0" border="0" ` +
    `style="width:100%;max-width:560px;border-collapse:separate;margin:4px 0 12px;">` +
    `<tr><td style="background:#ffffff;border:1px solid #e0e0e0;border-left:4px solid ${ACCENTS[card.accent]};` +
    `border-radius:6px;padding:14px 16px;font-family:${FONT};color:#242424;">` +
    `<div style="font-size:15px;line-height:20px;font-weight:600;color:#242424;margin:0 0 8px;">` +
    `${escapeHtml(card.title)}</div>` +
    (card.highlight
      ? `<div style="font-family:Consolas,'Courier New',monospace;font-size:28px;line-height:36px;font-weight:700;` +
        `letter-spacing:6px;color:#242424;margin:0 0 4px;">${escapeHtml(card.highlight)}</div>`
      : "") +
    (rows ? `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">${rows}</table>` : "") +
    button +
    steps +
    details +
    `</td></tr></table>`
  );
}

/**
 * A button (the card's button and the buttons of its steps), the same as in the link card (addin mailCard.ts): the
 * colour on a table cell, so classic Outlook (Word engine) shows a filled button, not blue text in a white box.
 */
function buttonLink(button: { label: string; url: string }): string {
  return (
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;">` +
    `<tr><td bgcolor="${BRAND}" style="background:${BRAND};border-radius:4px;mso-padding-alt:6px 14px;">` +
    `<a href="${escapeHtml(button.url)}" style="display:inline-block;color:#ffffff;font-family:${FONT};font-size:13px;` +
    `font-weight:600;line-height:18px;padding:6px 14px;text-decoration:none;white-space:nowrap;">${escapeHtml(button.label)}</a>` +
    `</td></tr></table>`
  );
}

/** Small grey line below everything (e.g. the plain link for clients that do not show buttons). */
export const smallPrint = (html: string) =>
  `<p style="margin:0 0 8px;font-family:${FONT};font-size:12px;line-height:16px;color:#8a8a8a;">${html}</p>`;

/**
 * The whole mail: `opening` (paragraphs above the card, HTML), the card, `after` (HTML below it) and the line saying
 * that CloudShare sent it automatically.
 */
export function mailDocument(parts: { opening: string; card: MailCard; after?: string }): string {
  return (
    parts.opening +
    mailCard(parts.card) +
    (parts.after ?? "") +
    `<p style="margin:16px 0 0;padding:10px 0 0;border-top:1px solid #e0e0e0;font-family:${FONT};font-size:12px;` +
    `line-height:16px;color:#8a8a8a;">${escapeHtml(t("mail.automatic"))}</p>`
  );
}
