/**
 * Mails to the sender about a delivery proof (proof.ts), in the layout of mailLayout.ts and the language of the sender's
 * add-in, sent with CloudShare's app token (application permission Mail.Send – no user token exists then):
 *
 *   download – on the first download of a file of the case, with the PDF attached (it replaces the download confirmation
 *              of notices.ts for such links, so there is one mail, not two);
 *   final    – when all links of the case ended (expired or revoked), with the complete PDF;
 *   locked   – the gateway refused code requests for a while: too many addresses that are no recipients (proofGateway.ts);
 *   reminder – nothing fetched yet a few days before the links end (proofWatcher.ts, PROOF_REMINDER_DAYS).
 */
import { appGraphToken } from "./auth.js";
import { sendMailAsApp } from "./graph.js";
import { formatDate, formatDateTime, inSenderLanguage, t } from "./i18n.js";
import { mailDocument, paragraph } from "./mailLayout.js";
import { addEvent, caseSummary, type ProofCase } from "./proof.js";
import { signedProofPdf } from "./proofPdf.js";

/** Subject of the mail the case belongs to (the last one recorded as sent). */
export const caseSubject = (c: ProofCase) =>
  String([...c.events].reverse().find((e) => e.type === "message.sent")?.data.subject ?? "");

const time = (ms: number) => formatDateTime(ms);

/** The signed PDF as an attachment ("Zustellnachweis-20261117-1146-7Q2F.pdf"). */
const pdfAttachment = async (c: ProofCase) => ({
  name: inSenderLanguage(c.senderLanguage, () => `${t("mail.proof.fileName")}-${c.id}.pdf`.replace(/[^\w.-]/g, "_")),
  contentType: "application/pdf",
  content: await signedProofPdf(c, c.senderLanguage),
});

/** The first download of `name` at `at`: a green card, the PDF attached. */
export async function sendDownloadProof(c: ProofCase, name: string, at: number): Promise<void> {
  const mail = inSenderLanguage(c.senderLanguage, () => {
    const subject = caseSubject(c) || t("mail.noSubject");
    return {
      subject: t("mail.proof.downloadMail.subject", { id: c.id, name }),
      html: mailDocument({
        opening: paragraph(t("mail.proof.downloadMail.text", { name, subject })),
        card: { accent: "success", title: t("mail.proof.downloadMail.title"), files: [name], details: t("mail.proof.downloadMail.details", { time: time(at), id: c.id }) },
      }),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html, [await pdfAttachment(c)]);
  addEvent(c.id, "proof.sent", { reason: "download" });
}

/** All links of the case ended: the complete PDF. */
export async function sendFinalProof(c: ProofCase): Promise<void> {
  const mail = inSenderLanguage(c.senderLanguage, () => {
    const subject = caseSubject(c) || t("mail.noSubject");
    const names = Object.values(c.links ?? {}).map((l) => l.name);
    return {
      subject: t("mail.proof.finalMail.subject", { id: c.id }),
      html: mailDocument({
        opening: paragraph(t("mail.proof.finalMail.text", { subject })),
        card: { accent: "brand", title: t("mail.proof.finalMail.title"), files: names, details: t("mail.proof.finalMail.details", { count: c.events.length, id: c.id }) },
      }),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html, [await pdfAttachment(c)]);
  addEvent(c.id, "proof.sent", { reason: "final" });
}

/** Nothing fetched yet and the links end on `endsOn` (YYYY-MM-DD): the sender may ask the recipients. */
export async function sendReminder(c: ProofCase, endsOn: string): Promise<void> {
  const mail = inSenderLanguage(c.senderLanguage, () => {
    const subject = caseSubject(c) || t("mail.noSubject");
    const date = formatDate(Date.parse(`${endsOn}T12:00:00Z`));
    const summary = caseSummary(c);
    const recipients = [...c.events].reverse().find((e) => e.type === "message.sent")?.data.to;
    const to = Array.isArray(recipients) ? recipients.join(", ") : "";
    return {
      subject: t("mail.proof.reminderMail.subject", { id: c.id }),
      html: mailDocument({
        opening: paragraph(t("mail.proof.reminderMail.text", { subject, date })),
        card: {
          accent: "brand",
          title: t(summary.state === "confirmed" ? "mail.proof.reminderMail.titleConfirmed" : "mail.proof.reminderMail.title"),
          files: Object.values(c.links ?? {}).map((l) => l.name),
          details: [to ? t("mail.proof.reminderMail.to", { to }) : "", t("mail.proof.reminderMail.details", { date, id: c.id })].filter(Boolean).join(" · "),
        },
        after: paragraph(t("mail.proof.reminderMail.advice")),
      }),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html);
  addEvent(c.id, "proof.reminded", { endsOn });
}

/** The gateway stopped taking code requests for a link: too many addresses that are no recipients. */
export async function sendLockNotice(c: ProofCase, linkName: string, attempts: Array<{ at: number; address: string; ip: string }>): Promise<void> {
  const mail = inSenderLanguage(c.senderLanguage, () => {
    const subject = caseSubject(c) || t("mail.noSubject");
    return {
      subject: t("mail.proof.lockMail.subject", { id: c.id }),
      html: mailDocument({
        // The attempts as plain lines (time · address · IP) – no file rows with type badges for them.
        opening: paragraph(t("mail.proof.lockMail.text", { subject })) +
          attempts.map((a) => paragraph(`${time(a.at)} · ${a.address} · ${a.ip}`)).join("") +
          paragraph(t("mail.proof.lockMail.advice")),
        card: {
          accent: "danger",
          title: t("mail.proof.lockMail.title"),
          files: [linkName],
          details: t("mail.proof.lockMail.details", { count: attempts.length, id: c.id }),
        },
      }),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html);
}
