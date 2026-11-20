/**
 * "Feedback senden" in the add-in: a bug report or an idea, mailed from the user's own mailbox (delegated Mail.Send –
 * the address of the organisation answers the user directly) to FEEDBACK_MAIL. Nothing goes anywhere public: whoever
 * reads it decides what to pass on (e.g. as a GitHub issue, without names, addresses and file names – the mail says
 * so). The technical details come from the pane (shown to the user first) and are limited to known fields.
 */
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { escapeHtml, mailDocument, paragraph, smallPrint, FONT } from "./mailLayout.js";
import { t } from "./i18n.js";

export type FeedbackKind = "bug" | "idea";

export interface Feedback {
  kind: FeedbackKind;
  text: string;
  /** Known fields only (DIAGNOSTICS), each cut to its length; absent when the user left them out. */
  diagnostics?: Record<string, string>;
}

export const FEEDBACK_MAX_LENGTH = 5000;
/** The details the pane may send (addin feedback.ts) – anything else is dropped. */
export const DIAGNOSTICS = ["outlook", "mailboxApi", "outlookLanguage", "addinLanguage", "window", "browser", "time", "errors"] as const;

/** The request body as a Feedback, or null when it is not one (no text, too long, unknown kind). */
export function parseFeedback(body: unknown): Feedback | null {
  if (!body || typeof body !== "object") return null;
  const { kind, text, diagnostics } = body as Record<string, unknown>;
  if (kind !== "bug" && kind !== "idea") return null;
  if (typeof text !== "string" || !text.trim() || text.length > FEEDBACK_MAX_LENGTH) return null;
  const feedback: Feedback = { kind, text: text.trim() };
  if (diagnostics && typeof diagnostics === "object") {
    const known: Record<string, string> = {};
    for (const key of DIAGNOSTICS) {
      const value = (diagnostics as Record<string, unknown>)[key];
      if (typeof value === "string" && value.trim()) known[key] = value.trim().slice(0, key === "errors" ? 2000 : 500);
    }
    feedback.diagnostics = known;
  }
  return feedback;
}

let version: string | undefined;

/** The add-in's version: `<Version>` of the manifest the server serves ("?" without one, e.g. in tests). */
export function addinVersion(): string {
  if (version === undefined) {
    const manifest = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../public/manifest.xml");
    version = existsSync(manifest) ? (/<Version>([^<]+)<\/Version>/.exec(readFileSync(manifest, "utf8"))?.[1] ?? "?") : "?";
  }
  return version;
}

/** Plain text with its line breaks, escaped – one paragraph per block separated by an empty line. */
const textBlocks = (text: string) =>
  text
    .split(/\n\s*\n/)
    .map((block) => `<p style="margin:0 0 12px;font-family:${FONT};font-size:14px;line-height:20px;color:#242424;">` +
      `${escapeHtml(block.trim()).replace(/\n/g, "<br>")}</p>`)
    .join("");

/** Subject and HTML of the mail to FEEDBACK_MAIL (in the organisation's language – the caller sets it). */
export function feedbackMail(feedback: Feedback, sender: { name: string; address: string }): { subject: string; html: string } {
  const bug = feedback.kind === "bug";
  const details = feedback.diagnostics
    ? `<table role="presentation" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;margin:0 0 12px;">` +
      Object.entries({ addinVersion: addinVersion(), ...feedback.diagnostics })
        .map(([key, value]) =>
          `<tr><td valign="top" style="padding:2px 12px 2px 0;font-family:${FONT};font-size:12px;line-height:16px;color:#616161;white-space:nowrap;">` +
          `${escapeHtml(t(`mail.feedbackDetail.${key}`))}</td>` +
          `<td valign="top" style="padding:2px 0;font-family:Consolas,'Courier New',monospace;font-size:12px;line-height:16px;color:#242424;word-break:break-all;">` +
          `${escapeHtml(value).replace(/\n/g, "<br>")}</td></tr>`)
        .join("") +
      `</table>`
    : paragraph(t("mail.feedbackNoDetails"));
  return {
    subject: t(bug ? "mail.feedbackSubjectBug" : "mail.feedbackSubjectIdea", { name: sender.name }),
    html: mailDocument({
      opening: paragraph(t(bug ? "mail.feedbackIntroBug" : "mail.feedbackIntroIdea", { name: sender.name, address: sender.address })),
      card: {
        accent: bug ? "danger" : "brand",
        title: t(bug ? "mail.feedbackTitleBug" : "mail.feedbackTitleIdea"),
        files: [],
        details: t("mail.feedbackReply", { name: sender.name }),
      },
      after:
        textBlocks(feedback.text) +
        `<p style="margin:16px 0 6px;font-family:${FONT};font-size:14px;line-height:20px;font-weight:600;color:#242424;">` +
        `${escapeHtml(t("mail.feedbackDetails"))}</p>` +
        details +
        smallPrint(escapeHtml(t("mail.feedbackPublic"))),
    }),
  };
}
