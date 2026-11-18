import { HttpError } from "./auth.js";
import { config } from "./config.js";
import { t } from "./i18n.js";

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
  if (!res.ok) throw new HttpError(502, t("server.profileUnavailable", { status: res.status }));
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
  /** Graph ID of the sent mail (immutable: it survives moving the mail) – the pane opens the mail with it. */
  id?: string;
  /** When it was sent (ISO) – shown in "Verlauf" with the recipients. */
  sentAt?: string;
  /** Internet Message-ID ("<…@…>") – names the mail in Exchange's message trace ("Zustellnachweis"). */
  messageId?: string;
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
    throw new HttpError(502, `Graph sendMail failed (HTTP ${res.status}): ${text.slice(0, 300)}`);
  }
}

/**
 * Sends a mail as `userId` (oid or UPN) with CloudShare's own app token – the download confirmation to the sender.
 * Not saved in "Sent Items"; needs the application permission Mail.Send.
 */
/** A file sent with a mail (the PDF of a delivery proof). */
export interface MailAttachment {
  name: string;
  contentType: string;
  content: Buffer;
}

export async function sendMailAsApp(
  appToken: string,
  userId: string,
  to: string,
  subject: string,
  html: string,
  attachments: MailAttachment[] = [],
) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(userId)}/sendMail`, {
    method: "POST",
    headers: { Authorization: `Bearer ${appToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: "HTML", content: html },
        toRecipients: toGraph([to]),
        ...(attachments.length
          ? {
              attachments: attachments.map((a) => ({
                "@odata.type": "#microsoft.graph.fileAttachment",
                name: a.name,
                contentType: a.contentType,
                contentBytes: a.content.toString("base64"),
              })),
            }
          : {}),
      },
      saveToSentItems: false,
    }),
  });
  if (res.status !== 202) {
    const text = await res.text().catch(() => "");
    throw new HttpError(502, `Graph sendMail (app) failed (HTTP ${res.status}): ${text.slice(0, 300)}`);
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
    `$select=${encodeURIComponent("subject,from,toRecipients,ccRecipients,bccRecipients,sentDateTime,internetMessageId")}`,
    `$expand=${encodeURIComponent(`singleValueExtendedProperties($filter=id eq '${customPropertiesId()}')`)}`,
  ].join("&");
  const res = await fetch(`https://graph.microsoft.com/v1.0/${root}/mailFolders/sentitems/messages?${query}`, {
    // Immutable IDs: the ID stays the same when the user moves the mail out of "Sent Items".
    headers: { Authorization: `Bearer ${graphToken}`, Prefer: `IdType="ImmutableId"` },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph sent items (${mailbox ?? "own mailbox"}): HTTP ${res.status} ${text.slice(0, 300)}`);
  }
  const json = (await res.json()) as {
    value?: Array<{
      id?: string;
      sentDateTime?: string;
      internetMessageId?: string;
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
      id: message.id,
      sentAt: message.sentDateTime,
      messageId: message.internetMessageId,
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

/**
 * The EWS ID Outlook needs to open a mail (Office.js displayMessageForm), for the immutable Graph ID stored when the
 * mail was found in "Sent Items" ("Verlauf": click on the subject). Translated now, so it is the mail's current ID even
 * after it was moved. Null if Exchange does not know it (any more). Needs only User.Read.
 */
export async function ewsIdFor(graphToken: string, immutableId: string): Promise<string | null> {
  const res = await fetch("https://graph.microsoft.com/v1.0/me/translateExchangeIds", {
    method: "POST",
    headers: { Authorization: `Bearer ${graphToken}`, "Content-Type": "application/json" },
    body: JSON.stringify({ inputIds: [immutableId], sourceIdType: "restImmutableEntryId", targetIdType: "ewsId" }),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new HttpError(502, `Graph translateExchangeIds failed (HTTP ${res.status}): ${text.slice(0, 300)}`);
  }
  const json = (await res.json()) as { value?: Array<{ targetId?: string }> };
  return json.value?.[0]?.targetId || null;
}

/**
 * Whether the sent mail is still where the user keeps mails: "deleted" when it lies in "Deleted Items" or – after the
 * folder was emptied – in the hidden "Recoverable Items" (Exchange keeps it there for a while, and Outlook would still
 * open it by its ID), "gone" when Exchange no longer has it. Needs Mail.Read.
 */
export async function sentMailState(graphToken: string, immutableId: string): Promise<"ok" | "deleted" | "gone"> {
  const headers = { Authorization: `Bearer ${graphToken}`, Prefer: `IdType="ImmutableId"` };
  const get = (path: string) => fetch(`https://graph.microsoft.com/v1.0/me/${path}`, { headers });
  const res = await get(`messages/${encodeURIComponent(immutableId)}?$select=parentFolderId`);
  if (res.status === 404) return "gone";
  if (!res.ok) throw new Error(`Graph message: HTTP ${res.status}`);
  const { parentFolderId } = (await res.json()) as { parentFolderId?: string };
  for (const folder of ["deleteditems", "recoverableitemsdeletions"]) {
    const f = await get(`mailFolders/${folder}?$select=id`);
    if (f.ok && ((await f.json()) as { id?: string }).id === parentFolderId) return "deleted";
  }
  return "ok";
}

/**
 * Delegated permission for the sensitivity labels published to the user (Microsoft Purview) – optional: without admin
 * consent the add-in falls back to Outlook's own label catalog.
 */
export const GRAPH_LABEL_SCOPES = ["https://graph.microsoft.com/InformationProtectionPolicy.Read"];

export interface SensitivityLabel {
  /** GUID, lower case without braces. */
  id: string;
  name: string;
  /** Priority in the label policy (Purview "sensitivity"): higher is more sensitive. */
  rank: number;
  /** The label encrypts what it is applied to. */
  protection: boolean;
}

/**
 * The sensitivity labels the user may apply, with their names in `language` and their priority – for the label of a
 * shared file and which one is higher. Sublabels come with their parent's name ("Vertraulich / Alle Mitarbeiter").
 */
export async function fetchSensitivityLabels(graphToken: string, language?: string): Promise<SensitivityLabel[]> {
  const labels: SensitivityLabel[] = [];
  type GraphLabel = { id?: string; name?: string; sensitivity?: number; hasProtection?: boolean; isActive?: boolean; parent?: { name?: string } | null };
  let url: string | undefined = "https://graph.microsoft.com/v1.0/me/security/informationProtection/sensitivityLabels";
  for (let page = 0; url && page < 10; page++) {
    const res: Response = await fetch(url, {
      headers: { Authorization: `Bearer ${graphToken}`, ...(language ? { "Accept-Language": language } : {}) },
    });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new HttpError(502, `Graph sensitivityLabels failed (HTTP ${res.status}): ${text.slice(0, 300)}`);
    }
    const json = (await res.json()) as { value?: GraphLabel[]; "@odata.nextLink"?: string };
    for (const label of json.value ?? []) {
      if (!label.id || !label.name || label.isActive === false) continue;
      labels.push({
        id: label.id.replace(/[{}]/g, "").toLowerCase(),
        name: label.parent?.name ? `${label.parent.name} / ${label.name}` : label.name,
        rank: typeof label.sensitivity === "number" ? label.sensitivity : labels.length,
        protection: label.hasProtection === true,
      });
    }
    url = json["@odata.nextLink"];
  }
  return labels;
}
