/**
 * Mails to the sender about a delivery proof (proof.ts), in the layout of mailLayout.ts and the language of the sender's
 * add-in, sent with CloudShare's app token (application permission Mail.Send – no user token exists then):
 *
 *   download – on the first download of a file of the case, with the PDF attached (it replaces the download confirmation
 *              of notices.ts for such links, so there is one mail, not two);
 *   final    – when all links of the case ended (expired or revoked), with the complete PDF;
 *   locked   – the gateway refused code requests for a while: too many addresses that are no recipients (proofGateway.ts).
 */
import { appGraphToken } from "./auth.js";
import { sendMailAsApp } from "./graph.js";
import { formatDateTime, inSenderLanguage, t } from "./i18n.js";
import { mailDocument, paragraph } from "./mailLayout.js";
import { addEvent, type ProofCase } from "./proof.js";
import { proofPdf } from "./proofPdf.js";

/** Subject of the mail the case belongs to (the last one recorded as sent). */
export const caseSubject = (c: ProofCase) =>
  String([...c.events].reverse().find((e) => e.type === "message.sent")?.data.subject ?? "");

const time = (ms: number) => formatDateTime(ms);

/** The PDF as an attachment ("Zustellnachweis-20261117-1146-7Q2F.pdf"). */
const pdfAttachment = (c: ProofCase) => ({
  name: `${t("mail.proof.fileName")}-${c.id}.pdf`.replace(/[^\w.-]/g, "_"),
  contentType: "application/pdf",
  content: proofPdf(c, c.senderLanguage),
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
      attachment: pdfAttachment(c),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html, [mail.attachment]);
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
      attachment: pdfAttachment(c),
    };
  });
  await sendMailAsApp(await appGraphToken(), c.owner, c.ownerUpn, mail.subject, mail.html, [mail.attachment]);
  addEvent(c.id, "proof.sent", { reason: "final" });
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
