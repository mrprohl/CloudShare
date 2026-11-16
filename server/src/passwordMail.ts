/**
 * The mails about a password, in the common layout of mailLayout.ts:
 *   - the password mail to the recipients (texts from section "mail" of the language file, own intro from
 *     templates.json): a card with the files and the button to the password link (never the password itself);
 *   - the note to the sender that a password SMS failed: a card with the files and the number, no password, no link.
 */
import { t } from "./i18n.js";
import { escapeHtml, mailDocument, paragraph, smallPrint, TEXT_STYLE } from "./mailLayout.js";
import { renderHtml } from "./templates.js";

export { escapeHtml } from "./mailLayout.js";

export const accessText = (count?: number) => (count === undefined ? t("mail.accessLimited") : t("mail.access", { count }));

export interface PasswordMailContent {
  fileNames: string[];
  sendUrl: string;
  maxAccessCount?: number;
  /** Subject of the mail with the links ("" if unknown). */
  subject: string;
  /** Last day the password link works (formatted date); older jobs without it show `expireDays`. */
  expiresOn?: string;
  /** Days until the password link expires (jobs from before `expiresOn`). */
  expireDays: number;
  /**
   * Own intro text from templates.json (mail_password_intro: greeting and text above the card); without it the
   * built-in text of the language is used.
   */
  introTemplate?: string;
  senderName?: string;
  /** The sender's signature (HTML from signature.ts, already safe) – below the card, so recipients see who sent it. */
  signature?: string;
}

export function passwordMailHtml(mail: PasswordMailContent): string {
  const count = mail.fileNames.length;
  const opening = mail.introTemplate
    ? renderHtml(mail.introTemplate, { subject: mail.subject, sender_name: mail.senderName ?? "", count: String(count) }, TEXT_STYLE)
    : paragraph(t("mail.greeting")) +
      paragraph(mail.subject ? t("mail.introSubject", { count, subject: mail.subject }) : t("mail.intro", { count }));
  const expiry = mail.expiresOn ? t("mail.expiresOn", { date: mail.expiresOn }) : t("mail.expires", { count: mail.expireDays });
  return mailDocument({
    opening,
    card: {
      accent: "brand",
      title: t("mail.cardTitle", { count }),
      files: mail.fileNames,
      button: { label: t("mail.showPassword"), url: mail.sendUrl },
      details: [accessText(mail.maxAccessCount), expiry, t("mail.enterPassword")].join(" · "),
    },
    // The plain link as well – some mail clients block buttons or show them without colours.
    after:
      smallPrint(
        `${escapeHtml(t("mail.buttonFallback"))} <a href="${escapeHtml(mail.sendUrl)}" style="color:#0f6cbd;">` +
          `${escapeHtml(mail.sendUrl)}</a>`,
      ) + (mail.signature ? `<div style="${TEXT_STYLE}margin:16px 0 0;">${mail.signature}</div>` : ""),
  });
}

/**
 * Note to the sender that the password SMS of a step failed – without the password and without a password link:
 * nothing goes out another way on its own, the sender sends the password again in CloudShare. `reason`: why seven.io
 * refused the SMS; without it, seven.io reported later that it was not delivered.
 */
export function smsProblemMailHtml(note: { fileNames: string[]; subject: string; to: string; reason?: string }): string {
  const undelivered = note.reason === undefined;
  const subject = note.subject || t("mail.noSubject");
  return mailDocument({
    opening:
      paragraph(t("mail.greeting")) +
      paragraph(t(undelivered ? "mail.smsUndeliveredText" : "mail.smsFailedText", { subject })),
    card: {
      accent: "danger",
      title: t(undelivered ? "mail.smsUndeliveredTitle" : "mail.smsFailedTitle"),
      files: note.fileNames,
      details: [t("mail.smsTo", { to: note.to }), undelivered ? "" : t("mail.smsReason", { reason: note.reason! })]
        .filter(Boolean)
        .join(" · "),
    },
    after: paragraph(t("mail.smsFailedHint")),
  });
}
