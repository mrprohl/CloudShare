import { HttpError } from "./auth.js";
import { config } from "./config.js";

/**
 * Delegated Graph permissions: read "Sent Items" to detect the sent mail, then send the password mail.
 * The .Shared variants cover mails sent from shared mailboxes / on behalf of others – they only work for
 * mailboxes the user already has access and send rights to in Exchange.
 */
export const GRAPH_SCOPES = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Mail.Read.Shared",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/Mail.Send.Shared",
];

/**
 * Name of the Office.js custom property that links a draft – and its copy in "Sent Items" – to a password job.
 * Must match JOB_PROPERTY in addin/src/shared/office.ts.
 */
export const JOB_PROPERTY = "cloudshareJob";

/**
 * Office.js stores an add-in's custom properties as one JSON string in a MAPI named property
 * "cecp-<add-in ID>" of the PS_PUBLIC_STRINGS property set; it travels with the item into "Sent Items".
 */
const customPropertiesId = () => `String {00020329-0000-0000-C000-000000000046} Name cecp-${config.addinId}`;

/** Delegated permission for the sender's profile (footer placeholders); granted by default in Entra. */
export const GRAPH_PROFILE_SCOPES = ["https://graph.microsoft.com/User.Read"];

/** Profile fields available as footer placeholders ({displayName}, {jobTitle}, …). */
export interface Profile {
  displayName: string;
  givenName: string;
  surname: string;
  jobTitle: string;
  department: string;
  company: string;
  email: string;
  phone: string;
  mobile: string;
  office: string;
}

/** The signed-in user's profile from Entra ID / Microsoft 365. */
export async function fetchProfile(graphToken: string): Promise<Profile> {
  const select = "displayName,givenName,surname,jobTitle,department,companyName,mail,userPrincipalName,businessPhones,mobilePhone,officeLocation";
  const res = await fetch(`https://graph.microsoft.com/v1.0/me?$select=${select}`, {
    headers: { Authorization: `Bearer ${graphToken}` },
  });
  if (!res.ok) throw new HttpError(502, `Profil konnte nicht gelesen werden (HTTP ${res.status})`);
  const me = (await res.json()) as Record<string, unknown>;
  const text = (v: unknown) => (typeof v === "string" ? v.trim() : "");
  return {
    displayName: text(me.displayName),
    givenName: text(me.givenName),
    surname: text(me.surname),
    jobTitle: text(me.jobTitle),
    department: text(me.department),
    company: text(me.companyName),
    email: text(me.mail) || text(me.userPrincipalName),
    phone: Array.isArray(me.businessPhones) ? text(me.businessPhones[0]) : "",
    mobile: text(me.mobilePhone),
    office: text(me.officeLocation),
  };
}

export interface MailRecipients {
  to: string[];
  cc: string[];
  bcc: string[];
}

export interface SentMessage {
  subject: string;
  /** Address in the "From" field – the user's own or a shared mailbox (send as / on behalf). */
  from: string | null;
  recipients: MailRecipients;
}

type GraphRecipient = { emailAddress?: { address?: string } };

const toGraph = (addresses: string[]) => addresses.map((address) => ({ emailAddress: { address } }));
const fromGraph = (list: GraphRecipient[] | undefined) =>
  (list ?? []).map((r) => r.emailAddress?.address).filter((a): a is string => Boolean(a));

/**
 * Sends a mail with the signed-in user's delegated token. With `from` set to another mailbox it is sent as /
 * on behalf of that mailbox – Exchange applies the user's Send As / Send on Behalf rights.
 */
export async function sendMail(
  graphToken: string,
  recipients: MailRecipients,
  subject: string,
  html: string,
  from?: string | null,
) {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
    method: "POST",
    headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: "HTML", content: html },
        toRecipients: toGraph(recipients.to),
        ccRecipients: toGraph(recipients.cc),
        bccRecipients: toGraph(recipients.bcc),
        ...(from ? { from: { emailAddress: { address: from } } } : {}),
      },
      saveToSentItems: true,
    }),
  });
  if (res.status !== 202) {
    const text = await res.text().catch(() => "");
    throw new HttpError(502, `Graph sendMail fehlgeschlagen (HTTP ${res.status}): ${text.slice(0, 300)}`);
  }
}

/**
 * Looks in "Sent Items" for the mail tagged with the job ID (sent since `since`) – in the user's own mailbox, or
 * in `mailbox` (shared mailbox address) if given. Returns its actual sender, subject and recipients, or null if it
 * has not been sent (yet).
 */
export async function findSentMessage(
  graphToken: string,
  jobId: string,
  since: number,
  mailbox?: string,
): Promise<SentMessage | null> {
  const root = mailbox ? `users/${encodeURIComponent(mailbox)}` : "me";
  const query = [
    `$filter=${encodeURIComponent(`sentDateTime ge ${new Date(since).toISOString()}`)}`,
    `$orderby=${encodeURIComponent("sentDateTime desc")}`,
    `$top=25`,
    `$select=${encodeURIComponent("subject,from,toRecipients,ccRecipients,bccRecipients,sentDateTime")}`,
    `$expand=${encodeURIComponent(`singleValueExtendedProperties($filter=id eq '${customPropertiesId()}')`)}`,
  ].join("&");
  const res = await fetch(`https://graph.microsoft.com/v1.0/${root}/mailFolders/sentitems/messages?${query}`, {
    headers: { Authorization: `Bearer ${graphToken}` },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph sent items (${mailbox ?? "own mailbox"}): HTTP ${res.status} ${text.slice(0, 300)}`);
  }
  const json = (await res.json()) as {
    value?: Array<{
      subject?: string;
      from?: GraphRecipient;
      toRecipients?: GraphRecipient[];
      ccRecipients?: GraphRecipient[];
      bccRecipients?: GraphRecipient[];
      singleValueExtendedProperties?: Array<{ id: string; value: string }>;
    }>;
  };
  for (const message of json.value ?? []) {
    const raw = message.singleValueExtendedProperties?.[0]?.value;
    if (!raw) continue;
    try {
      if ((JSON.parse(raw) as Record<string, unknown>)[JOB_PROPERTY] !== jobId) continue;
    } catch {
      continue;
    }
    return {
      subject: message.subject ?? "",
      from: message.from?.emailAddress?.address ?? null,
      recipients: {
        to: fromGraph(message.toRecipients),
        cc: fromGraph(message.ccRecipients),
        bcc: fromGraph(message.bccRecipients),
      },
    };
  }
  return null;
}
