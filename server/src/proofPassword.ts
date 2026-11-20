/**
 * "Zustellnachweis": the password of a step whose links all lead through the gateway is not sent with the mail
 * (jobs.ts holds it back, PasswordRef.deferred). It goes out only after a recipient confirmed their address with the
 * code (proofGateway.ts), so recipients do not get a link mail and a password mail at the same time:
 *
 *   recipients of the mail – the password mail to the confirmed address only (every recipient confirms for themselves);
 *   another address        – the password mail there, once, after the first confirmation;
 *   SMS                    – the SMS once, after the first confirmation.
 *
 * Like "send again" (app.ts): the password is read from the step's Send, a password mail gets a new Send with it
 * (deleted with the links), the SMS carries the password itself. Sent with CloudShare's app token as the sender's
 * mailbox; every delivery is recorded (recordDelivery → the case).
 */
import { appGraphToken } from "./auth.js";
import { config } from "./config.js";
import { recordDelivery } from "./deliveries.js";
import { sendMailAsApp } from "./graph.js";
import { formatDate, inRecipientLanguage, t } from "./i18n.js";
import { takeSmsQuota } from "./jobs.js";
import { passwordMailHtml } from "./passwordMail.js";
import { passwordRefFor, updatePasswordRef, type PasswordRef } from "./passwordRefs.js";
import { gatewayUrl, saveCase, type ProofCase } from "./proof.js";
import { maskPhone, sendSms, smsEnabled } from "./sms.js";
import { defaultTemplates, render, smsFiles, templatesFor } from "./templates.js";
import { createPasswordSend, readSendPassword } from "./vaultwarden.js";

/** What went out (now or before) for the page after the confirmation. */
export interface ReleasedPassword {
  /** The password mail to the confirmed address, with its subject. */
  mailTo?: string;
  subject?: string;
  /** Other addresses that got (or get) it, and the masked SMS number. */
  others?: string[];
  sms?: string;
}

interface Released {
  mailed: string[];
  others?: boolean;
  sms?: boolean;
}

/** One password mail to one address: a new Send with the step's password, in the recipients' language. */
async function mailPassword(c: ProofCase, ref: PasswordRef, password: string, to: string, kind?: "file" | "request", token = ""): Promise<string> {
  const send = await createPasswordSend(`CloudShare: ${ref.fileNames[0]}`, password, ref.maxAccessCount, config.sendExpireDays, ref.keepUntil);
  await updatePasswordRef(ref.sendId, { resentSend: send.id }); // deleted with the links
  const mail = inRecipientLanguage(ref.language, () => {
    const templates = templatesFor({ upn: ref.ownerUpn, department: ref.department });
    const customIntro = templates.mail_password_intro !== defaultTemplates().mail_password_intro ? templates.mail_password_intro : undefined;
    return {
      subject: render(templates.mail_password_subject, {
        files: t("mail.files", { count: ref.fileNames.length, name: ref.fileNames[0] }),
        subject: ref.subject || ref.fileNames[0],
        sender_name: ref.senderName ?? "",
      }),
      html: passwordMailHtml({
        fileNames: ref.fileNames, sendUrl: send.url, maxAccessCount: ref.maxAccessCount, subject: ref.subject,
        expireDays: config.sendExpireDays, introTemplate: customIntro, senderName: ref.senderName || ref.ownerUpn,
        expiresOn: formatDate(send.expiresAt), signature: ref.signature, trackingId: c.id,
        // Right after the confirmation: the mail says so and how to go on.
        afterConfirmation: kind ? { kind, url: gatewayUrl(token) } : undefined,
      }),
    };
  });
  // As the mailbox the mail came from (a shared mailbox: Exchange applies the app's access to it).
  await sendMailAsApp(await appGraphToken(), ref.from ?? c.owner, to, mail.subject, mail.html);
  await recordDelivery(ref.tokens, { owner: c.owner, channel: "mail", to, at: Date.now() });
  return mail.subject;
}

/**
 * The addresses that confirmed themselves with the code for one of `tokens` (lower case, in the order they did) –
 * "Kennwort erneut senden" of a held-back step may only reach them (app.ts), never someone who has not confirmed.
 */
export function verifiedAddresses(c: ProofCase | undefined, tokens: string[]): string[] {
  const found: string[] = [];
  for (const event of c?.events ?? []) {
    if (event.type !== "recipient.verified") continue;
    const links = Array.isArray(event.data.links) ? (event.data.links as unknown[]) : [];
    const address = typeof event.data.address === "string" ? event.data.address.toLowerCase() : "";
    if (address && links.some((link) => tokens.includes(String(link))) && !found.includes(address)) found.push(address);
  }
  return found;
}

/**
 * After `address` confirmed itself for the link `token`: sends what is still held back for it. Null when nothing was
 * held back (link without password, or a mail from before) – the gateway then goes straight on to the link.
 */
export async function releasePassword(c: ProofCase, token: string, address: string): Promise<ReleasedPassword | null> {
  const ref = passwordRefFor(token, c.owner);
  if (!ref?.deferred) return null;
  const done: Released = ((c.released ??= {})[ref.sendId] ??= { mailed: [] });
  const result: ReleasedPassword = {};
  const kind = c.links?.[token]?.kind === "request" ? "request" : "file";
  let password: string | undefined;
  const readPassword = async () => (password ??= await readSendPassword(ref.sendId));

  if (ref.toRecipients) {
    result.mailTo = address;
    if (!done.mailed.includes(address)) {
      result.subject = await mailPassword(c, ref, await readPassword(), address, kind, token);
      done.mailed.push(address);
      saveCase(c.id);
    }
  }
  if (ref.otherTo?.length) {
    result.others = ref.otherTo;
    if (!done.others) {
      // Not the one who confirmed: the usual password mail.
      for (const to of ref.otherTo) await mailPassword(c, ref, await readPassword(), to);
      done.others = true;
      saveCase(c.id);
    }
  }
  if (ref.sms) {
    result.sms = maskPhone(ref.sms);
    if (!done.sms && smsEnabled() && takeSmsQuota(c.owner)) {
      const pw = await readPassword();
      const text = inRecipientLanguage(ref.language, () =>
        render(templatesFor({ upn: ref.ownerUpn, department: ref.department }).sms_template, {
          password: pw, sender_name: ref.senderName ?? "", files: smsFiles(ref.fileNames),
        }));
      const messageId = await sendSms(ref.sms, text);
      await recordDelivery(ref.tokens, { owner: c.owner, channel: "sms", to: maskPhone(ref.sms), at: Date.now(), messageId, status: "submitted" });
      done.sms = true;
      saveCase(c.id);
    }
  }
  return result;
}
