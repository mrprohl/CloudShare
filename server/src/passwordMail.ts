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
  /** "Zustellnachweis": the tracking ID of the mail's case, named in the card. */
  trackingId?: string;
  /**
   * "Zustellnachweis": sent right after the recipient confirmed their address – the mail says so and lists the next
   * steps back on the page ("Zum Download" / "Zum Hochladen"). Own wording, not the template's intro.
   */
  afterConfirmation?: {
    kind: "file" | "request";
    /** The gateway of the link – a button "Zum Download" goes there. */
    url: string;
    /** Until when the link works ("04.12.2026, 23:59 Uhr") – a grey note under step 2. */
    validUntil?: string;
  };
}

export function passwordMailHtml(mail: PasswordMailContent): string {
  const count = mail.fileNames.length;
  const opening = mail.afterConfirmation
    ? paragraph(t("mail.greeting")) +
      paragraph(t("mail.confirmedIntro", { count, name: mail.fileNames[0], subject: mail.subject || mail.fileNames[0], sender: mail.senderName || "" }))
    : mail.introTemplate
    ? renderHtml(mail.introTemplate, { subject: mail.subject, sender_name: mail.senderName ?? "", count: String(count) }, TEXT_STYLE)
    : paragraph(t("mail.greeting")) +
      paragraph(mail.subject ? t("mail.introSubject", { count, subject: mail.subject }) : t("mail.intro", { count }));
  const expiry = mail.expiresOn ? t("mail.expiresOn", { date: mail.expiresOn }) : t("mail.expires", { count: mail.expireDays });
  // How often and how long the button works, in a sentence – the link stops working after that, and recipients should
  // keep the password until they are done. Older jobs without the date or the number: the short form.
  const limits = mail.expiresOn && mail.maxAccessCount
    ? t("mail.limitNote", { count: mail.maxAccessCount, date: mail.expiresOn })
    : [accessText(mail.maxAccessCount), expiry].filter(Boolean).join(" · ");
  const signature = mail.signature ? `<div style="margin:20px 0 0;">${mail.signature}</div>` : "";
  const tracking = mail.trackingId ? t("mail.trackingId", { id: mail.trackingId }) : "";
  if (mail.afterConfirmation) {
    // Right after the confirmation: two numbered steps in the card – first the password, then the download. Each has
    // its button and its address as plain text. The download goes to the gateway, not to Nextcloud: the confirmed
    // browser goes straight on, another one or a forwarded mail is asked for the code again – every opening is recorded.
    const upload = mail.afterConfirmation.kind === "request";
    const passwordLimits = mail.expiresOn && mail.maxAccessCount
      ? t("mail.confirmedLimit", { count: mail.maxAccessCount, date: mail.expiresOn })
      : [accessText(mail.maxAccessCount), expiry].filter(Boolean).join(" · ");
    return mailDocument({
      opening,
      card: {
        accent: "brand",
        title: t(upload ? "mail.confirmedTitleUpload" : "mail.confirmedTitleDownload"),
        files: mail.fileNames,
        steps: [
          // How often and how long the password link works – and, when it has run out, a new one through the gateway.
          { text: t("mail.confirmedStepPassword"), button: { label: t("mail.showPassword"), url: mail.sendUrl }, note: passwordLimits || undefined,
            noteLink: { label: t("mail.confirmedPasswordAgain"), url: `${mail.afterConfirmation.url}/password` } },
          {
            text: t(upload ? "mail.confirmedStepUpload" : "mail.confirmedStepDownload", { count }),
            button: { label: t(upload ? "mail.proof.gateway.upload" : "mail.proof.gateway.download"), url: mail.afterConfirmation.url },
            note: mail.afterConfirmation.validUntil ? t("mail.confirmedValidUntil", { date: mail.afterConfirmation.validUntil }) : undefined,
          },
        ],
        details: tracking,
      },
      after: signature,
    });
  }
  return mailDocument({
    opening,
    card: {
      accent: "brand",
      title: t("mail.cardTitle", { count }),
      files: mail.fileNames,
      button: { label: t("mail.showPassword"), url: mail.sendUrl },
      details: [limits, t("mail.enterPassword", { count }), tracking].filter(Boolean).join(" · "),
    },
    // The plain link as well – some mail clients block buttons or show them without colours.
    after:
      smallPrint(
        `${escapeHtml(t("mail.buttonFallback"))} <a href="${escapeHtml(mail.sendUrl)}" style="color:#0f6cbd;">` +
          `${escapeHtml(mail.sendUrl)}</a>`,
      ) + signature,
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
