/**
 * Text templates for what CloudShare sends on the sender's behalf: the SMS with the password, the password mail and
 * the download confirmation. Optional file TEMPLATES_FILE (JSON):
 *
 *   { "templates": { "default": {…}, "groups": { "<department>": {…} }, "users": { "<upn>": {…} } } }
 *
 * The first match wins per field (user → department of the sender → default → built-in texts of the language), so
 * every level may override single fields only. Placeholders are {name}; which ones a field may use is fixed below. Problems are logged when the
 * file is loaded – a broken file never stops CloudShare, it falls back to the built-in texts.
 */
import { readFileSync } from "node:fs";
import { config } from "./config.js";
import { texts } from "./i18n.js";

export const TEMPLATE_FIELDS = {
  /** SMS with the password in plain text – {password} required; never the link or a file name. */
  sms_template: ["password", "sender_name"],
  /** Subject of the password mail. */
  mail_password_subject: ["files", "subject", "sender_name"],
  /** Text above the card of the password mail (paragraphs separated by an empty line). */
  mail_password_intro: ["subject", "sender_name", "count"],
  /** Download confirmation to the sender. */
  download_notice_subject: ["filename", "subject"],
  download_notice_body: ["filename", "subject", "downloaded_at", "expires", "recipients"],
} as const;

export type TemplateField = keyof typeof TEMPLATE_FIELDS;
export type Templates = Record<TemplateField, string>;

/** Built-in texts: section "templates" of the language file (LANGUAGE). The SMS should use GSM characters only. */
export const defaultTemplates = (): Templates =>
  Object.fromEntries(Object.keys(TEMPLATE_FIELDS).map((field) => [field, texts("templates")[field] ?? ""])) as Templates;

type Level = Partial<Templates>;
interface TemplateFile {
  default: Level;
  groups: Record<string, Level>;
  users: Record<string, Level>;
}

/** GSM 03.38 default alphabet (plus the extension table) – other characters turn an SMS into UCS-2 (70 characters). */
const GSM =
  "@£$¥èéùìòÇ\nØø\rÅåΔ_ΦΓΛΩΠΨΣΘΞÆæßÉ !\"#¤%&'()*+,-./0123456789:;<=>?¡ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÑÜ§¿abcdefghijklmnopqrstuvwxyzäöñüà" +
  "^{}\\[~]|€";
export const isGsm = (text: string) => [...text].every((c) => GSM.includes(c));

/** Length of an SMS text in GSM units (characters of the extension table count twice). */
export const gsmLength = (text: string) => [...text].reduce((n, c) => n + ("^{}\\[~]|€".includes(c) ? 2 : 1), 0);

const placeholders = (text: string) => [...text.matchAll(/\{([a-z_]+)\}/g)].map((m) => m[1]);

/** Checks one level of the file; drops fields that would leak something (e.g. {password} in a mail). */
function validateLevel(level: unknown, where: string, warn: (msg: string) => void): Level {
  if (!level || typeof level !== "object") return {};
  const result: Level = {};
  for (const [field, value] of Object.entries(level as Record<string, unknown>)) {
    if (!(field in TEMPLATE_FIELDS)) {
      warn(`templates: unknown field "${field}" in ${where} – ignored`);
      continue;
    }
    if (typeof value !== "string" || !value.trim()) {
      warn(`templates: ${where}.${field} is not a text – ignored`);
      continue;
    }
    const allowed = TEMPLATE_FIELDS[field as TemplateField] as readonly string[];
    const used = placeholders(value);
    if (field !== "sms_template" && used.includes("password")) {
      warn(`templates: ${where}.${field} must not contain {password} (only the SMS may) – ignored`);
      continue;
    }
    if (field === "sms_template") {
      if (!used.includes("password")) {
        warn(`templates: ${where}.sms_template needs {password} – ignored`);
        continue;
      }
      if (used.some((p) => p === "link" || p === "filename")) {
        warn(`templates: ${where}.sms_template must not contain the link or a file name – ignored`);
        continue;
      }
      if (!isGsm(value)) warn(`templates: ${where}.sms_template has non-GSM characters – the SMS is limited to 70 characters`);
      // 12 characters for the password, about 25 for a sender name.
      const estimate = gsmLength(value.replace("{password}", "x".repeat(12)).replace("{sender_name}", "x".repeat(25)));
      if (estimate > 160) warn(`templates: ${where}.sms_template is about ${estimate} characters – billed as several SMS`);
    }
    const unknown = used.filter((p) => !allowed.includes(p));
    if (unknown.length) warn(`templates: ${where}.${field} uses unknown placeholders {${unknown.join("}, {")}} – left as is`);
    result[field as TemplateField] = value;
  }
  return result;
}

/** Reads and checks a template file; missing file → built-in texts only. Exported for tests. */
export function loadTemplates(file: string, warn: (msg: string) => void = console.warn): TemplateFile {
  let raw: string;
  try {
    raw = readFileSync(file, "utf8");
  } catch (err) {
    if ((err as NodeJS.ErrnoException).code !== "ENOENT") warn(`templates: ${file} cannot be read – built-in texts used`);
    return { default: {}, groups: {}, users: {} };
  }
  let parsed: { templates?: { default?: unknown; groups?: unknown; users?: unknown } };
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    warn(`templates: ${file} is not valid JSON (${(err as Error).message}) – built-in texts used`);
    return { default: {}, groups: {}, users: {} };
  }
  const t = parsed.templates ?? {};
  const map = (value: unknown, where: string) =>
    Object.fromEntries(
      Object.entries(value && typeof value === "object" ? (value as Record<string, unknown>) : {}).map(([key, level]) => [
        key.toLowerCase(),
        validateLevel(level, `${where}["${key}"]`, warn),
      ]),
    );
  return { default: validateLevel(t.default, "default", warn), groups: map(t.groups, "groups"), users: map(t.users, "users") };
}

let loaded: TemplateFile | undefined;

/** The texts for a sender: user (UPN / mail address) → department → default → built-in. */
export function templatesFor(sender: { upn?: string; department?: string }, file: TemplateFile = (loaded ??= loadTemplates(config.templatesFile))): Templates {
  return {
    ...defaultTemplates(),
    ...file.default,
    ...(sender.department ? file.groups[sender.department.toLowerCase()] : undefined),
    ...(sender.upn ? file.users[sender.upn.toLowerCase()] : undefined),
  };
}

/** Fills {name} placeholders; unknown ones stay visible so a typo is noticed. */
export const render = (text: string, values: Record<string, string>) =>
  text.replace(/\{([a-z_]+)\}/g, (match, key: string) => (key in values ? values[key] : match));

const escapeHtml = (s: string) =>
  s.replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]!);

/** Plain template text → HTML paragraphs (empty line = new paragraph, line break = <br>), values escaped. */
export function renderHtml(text: string, values: Record<string, string>, style: string): string {
  return render(escapeHtml(text), Object.fromEntries(Object.entries(values).map(([k, v]) => [k, escapeHtml(v)])))
    .split(/\n{2,}/)
    .map((paragraph) => `<p style="${style}">${paragraph.split("\n").join("<br>")}</p>`)
    .join("");
}
