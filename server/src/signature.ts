/**
 * Signature of the password mails, so recipients see who the password comes from: the sender's own text from the
 * settings, else MAIL_FOOTER – both filled from the sender's Microsoft 365 profile. Plain text only (no markup, every
 * character escaped); an empty line starts a new paragraph. Shown like the text of the mail ("Guten Tag, …").
 */
import { config } from "./config.js";
import type { Profile } from "./graph.js";
import { escapeHtml, TEXT_STYLE } from "./mailLayout.js";

/** Own signature of the settings (server/src/settings.ts): plain text with profile placeholders. */
export interface Signature {
  format: "text";
  content: string;
}

const SEPARATOR = " · ";

/**
 * One paragraph of a text template with profile placeholders ({displayName}, {jobTitle}, {phone} …): a line break
 * starts a new line, a " · " segment whose placeholders are all empty is dropped with its fixed text ("Tel. {phone}"
 * without a number), empty lines are dropped. Same rules as MAIL_FOOTER always had.
 */
export function fillTemplate(template: string, profile: Partial<Profile>): string {
  const values = profile as Record<string, string | undefined>;
  return template
    .split(/\r?\n|\\n/)
    .map((line) =>
      line
        .split(SEPARATOR)
        .map((segment) => {
          let placeholders = 0;
          let filled = 0;
          const text = segment.replace(/\{(\w+)\}/g, (_match, key: string) => {
            placeholders++;
            const value = (values[key] ?? "").trim();
            if (value) filled++;
            return value;
          });
          return placeholders > 0 && filled === 0 ? "" : text.trim();
        })
        .filter(Boolean)
        .join(SEPARATOR),
    )
    .filter(Boolean)
    .join("\n");
}

/**
 * The paragraphs of a signature: an empty line (also "\n\n" written literally, as in the env file) separates them;
 * each is filled from `profile`, paragraphs left empty are dropped.
 */
export function signatureParagraphs(template: string, profile: Partial<Profile>): string[] {
  return template
    .replace(/\\n/g, "\n")
    .split(/\r?\n[ \t]*\r?\n/)
    .map((paragraph) => fillTemplate(paragraph, profile))
    .filter(Boolean);
}

/**
 * HTML of the signature for a password mail: the own text or MAIL_FOOTER as paragraphs in the style of the mail
 * text, lines within a paragraph as line breaks. "" if there is none.
 */
export function signatureHtml(own: Signature | null | undefined, profile: Partial<Profile>): string {
  const template = own?.format === "text" ? own.content : config.mailFooter;
  return signatureParagraphs(template, profile)
    .map((paragraph) => `<p style="${TEXT_STYLE}">${escapeHtml(paragraph).split("\n").join("<br>")}</p>`)
    .join("");
}
