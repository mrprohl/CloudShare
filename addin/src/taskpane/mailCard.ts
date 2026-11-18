/**
 * HTML cards inserted into the mail body (seen by recipients – texts from the language file, keys "card.*").
 * Built for mail clients: tables and inline styles only, no classes, no external images. Classic Outlook (Word
 * engine) ignores rounded corners – the card then degrades to a plain bordered box.
 *
 * One layout for every kind (files, confidential message, upload link): title, a short sentence, the button on the
 * left, below it the details in grey (files: type, name, size – one block per file, each with its own button), and at
 * the bottom the line "Gültig bis … · Das Kennwort …".
 */

import { t } from "../shared/i18n";

export interface CardFile {
  name: string;
  /** Bytes; null if unknown (links inserted again from "Verlauf") – the size is then left out. */
  size: number | null;
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

/** Small type badge before the file name ("PDF", "XLSX", "DATEI" without extension) in the colour of its kind. */
function typeBadge(name: string): string {
  const ext = name.includes(".") ? name.split(".").pop()!.toLowerCase() : "";
  const label = escapeHtml(ext ? ext.toUpperCase().slice(0, 4) : t("card.fileBadge"));
  const color = TYPE_COLORS[ext] ?? "#616161";
  return (
    `<td valign="middle" style="padding:6px 8px 0 0;">` +
    `<span style="display:inline-block;background:${color};border-radius:3px;color:#ffffff;font-family:${FONT};` +
    `font-size:10px;font-weight:700;line-height:16px;padding:0 5px;">${label}</span></td>`
  );
}

/**
 * The button, built to look the same in classic Outlook (Word engine): the colour on a table cell (bgcolor, padding via
 * mso-padding-alt) – an <a> with padding alone shows there as blue text in a white box.
 */
function button(url: string, text: string): string {
  return (
    `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:separate;">` +
    `<tr><td bgcolor="${BRAND}" style="background:${BRAND};border-radius:4px;mso-padding-alt:6px 14px;">` +
    `<a href="${escapeHtml(url)}" style="display:inline-block;color:#ffffff;font-family:${FONT};font-size:13px;` +
    `font-weight:600;line-height:18px;padding:6px 14px;text-decoration:none;white-space:nowrap;">${escapeHtml(text)}</a>` +
    `</td></tr></table>`
  );
}

/** The short sentence under the title. */
const lead = (text: string) =>
  `<div style="font-size:14px;line-height:20px;color:#424242;margin:0 0 12px;">${escapeHtml(text)}</div>`;

/**
 * Outer frame: white card, thin border, brand accent on the left. `title` and `note` are plain text. No signature:
 * the sender's own Outlook signature is in the message, and the password mail carries it for CloudShare.
 */
function frame(title: string, rows: string, note: string): string {
  return (
    `<table role="presentation" width="560" cellpadding="0" cellspacing="0" border="0" ` +
    `style="width:100%;max-width:560px;border-collapse:separate;margin:12px 0;">` +
    `<tr><td style="background:#ffffff;border:1px solid #e0e0e0;border-left:4px solid ${BRAND};border-radius:6px;` +
    `padding:14px 16px;font-family:${FONT};color:#242424;">` +
    `<div style="font-size:15px;line-height:20px;font-weight:600;color:#242424;margin:0 0 8px;">${escapeHtml(title)}</div>` +
    rows +
    `<div style="font-size:12px;line-height:16px;color:#616161;margin:10px 0 0;">${escapeHtml(note)}</div>` +
    `</td></tr></table>`
  );
}

/**
 * "Zustellnachweis": the tracking ID of the mail's case; `gateway` – the link leads to the page where the recipient
 * confirms their address first (files, upload links), not for a confidential message.
 */
export interface CardProof {
  id: string;
  gateway: boolean;
}

/** What the recipients learn about the delivery proof: the tracking ID and that they confirm their address. */
const proofNote = (proof: CardProof | undefined) =>
  proof ? t(proof.gateway ? "card.proof" : "card.proofMessage", { id: proof.id }) : "";

/** "Gültig bis 28.11.2026 · Das Kennwort erhalten Sie in einer separaten Nachricht." */
/** With a delivery proof the password comes only after the recipient confirmed their address (the gateway). */
const passwordNote = (proof?: CardProof) => t(proof?.gateway ? "card.proofPassword" : "card.passwordNote");

const cardNote = (expireDate: string, withPassword: boolean, proof?: CardProof) =>
  [t("card.validUntil", { date: expireDate }), withPassword ? passwordNote(proof) : "", proofNote(proof)].filter(Boolean).join(" · ");

/** Card for shared files: per file a block – the download button, below it type, name and size in grey. */
export function filesCard(
  files: CardFile[],
  expireDate: string,
  formatSize: (bytes: number) => string,
  /** false for links "Ohne Kennwort" – then there is no password mail to point to. */
  withPassword = true,
  proof?: CardProof,
): string {
  // One table row per file: classic Outlook (Word engine) ignores margins and padding of <div>, but keeps them on table
  // cells – so the space between the blocks, the line between them and the gap under the button are cell paddings.
  const last = files.length - 1;
  const blocks =
    `<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="width:100%;border-collapse:collapse;">` +
    files
      .map(
        (f, i) =>
          `<tr><td style="padding:${i ? 12 : 0}px 0 ${i < last ? 12 : 0}px 0;${i ? "border-top:1px solid #f0f0f0;" : ""}">` +
          button(f.url, t("card.download")) +
          `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;">` +
          `<tr>${typeBadge(f.name)}` +
          `<td valign="middle" style="padding:6px 0 0 0;font-family:${FONT};font-size:13px;line-height:18px;color:#616161;word-break:break-all;">` +
          `${escapeHtml(f.name)}${f.size === null ? "" : ` · ${escapeHtml(formatSize(f.size))}`}</td>` +
          `</tr></table></td></tr>`,
      )
      .join("") +
    `</table>`;
  return frame(t("card.filesTitle", { count: files.length }), lead(t("card.filesText", { count: files.length })) + blocks,
    cardNote(expireDate, withPassword, proof));
}

/** Card for an upload link ("Dateien von Empfängern erhalten"). */
export function requestCard(url: string, expireDate: string, proof?: CardProof): string {
  return frame(t("card.requestTitle"), lead(t("card.requestText")) + button(url, t("card.upload")), cardNote(expireDate, true, proof));
}

/** Card for a confidential message: the text itself is not in the mail, only the password-protected link to it. */
export function messageCard(url: string, expireDate: string, proof?: CardProof): string {
  return frame(t("card.messageTitle"), lead(t("card.messageText")) + button(url, t("card.messageOpen")), cardNote(expireDate, true, proof));
}

/** Plain-text versions for mails composed in text format. */
export function filesText(
  files: CardFile[],
  expireDate: string,
  formatSize: (bytes: number) => string,
  withPassword = true,
  proof?: CardProof,
): string {
  const lines = files.map((f) => `- ${f.name}${f.size === null ? "" : ` (${formatSize(f.size)})`}: ${f.url}`);
  return `\n${t("card.filesPlain", { date: expireDate })}\n${lines.join("\n")}\n` + (withPassword ? `${passwordNote(proof)}\n` : "") +
    (proof ? `${proofNote(proof)}\n` : "");
}

export function messageText(url: string, expireDate: string, proof?: CardProof): string {
  return `\n${t("card.messagePlain", { date: expireDate })}\n${url}\n${t("card.passwordNote")}\n` + (proof ? `${proofNote(proof)}\n` : "");
}

export function requestText(url: string, expireDate: string, proof?: CardProof): string {
  return `\n${t("card.requestPlain", { date: expireDate })}\n${url}\n${passwordNote(proof)}\n` + (proof ? `${proofNote(proof)}\n` : "");
}
