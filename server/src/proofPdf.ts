/**
 * The "Zustellnachweis" as PDF (pdf.ts): the case with its tracking ID, the mail and its recipients, every event with
 * date and time and its hash, and whether the chain is intact. In the language of the sender's add-in.
 */
import { config } from "./config.js";
import { formatDate, formatDateTime, inSenderLanguage, locale, t } from "./i18n.js";
import { pdfDocument, type PdfLine } from "./pdf.js";
import { signPdf } from "./pdfSign.js";
import { caseSummary, verifyCase, type ProofCase, type ProofEvent } from "./proof.js";

const time = (ms: number) => formatDateTime(ms);

const isoTime = (iso: unknown) => (typeof iso === "string" && Number.isFinite(Date.parse(iso)) ? time(Date.parse(iso)) : "");

const size = (bytes: unknown) => {
  if (typeof bytes !== "number") return "";
  const units = ["B", "KB", "MB", "GB", "TB"];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  return `${value.toLocaleString(locale(), { maximumFractionDigits: unit ? 1 : 0 })} ${units[unit]} (${bytes.toLocaleString(locale())} B)`;
};

/** A way the password went: "recipients", "mail:<address>", "sms:<number>". */
const way = (value: unknown) => {
  const [channel, ...rest] = String(value).split(":");
  const to = rest.join(":");
  if (channel === "recipients") return t("mail.proof.way.recipients");
  return channel === "sms" ? t("mail.proof.way.sms", { to }) : to;
};

const list = (value: unknown) => (Array.isArray(value) ? value.filter((v) => typeof v === "string").join(", ") : "");
const text = (value: unknown) => (typeof value === "string" || typeof value === "number" ? String(value) : "");

/** Label and detail lines of an event. */
function describe(e: ProofEvent, c: ProofCase): { label: string; details: Array<[string, string]>; mono?: Array<[string, string]> } {
  const d = e.data;
  const names = list((Array.isArray(d.links) ? d.links : []).map((token: string) => c.links?.[token]?.name ?? token));
  const about: Array<[string, string]> = names ? [[t("mail.proof.field.links"), names]] : [];
  switch (e.type) {
    case "case.opened": {
      const client = (d.client ?? {}) as Record<string, unknown>;
      const clientText = [client.host, client.platform, client.version].filter(Boolean).join(" · ");
      return { label: t("mail.proof.event.caseOpened"), details: [[t("mail.proof.field.sender"), text(d.sender)], [t("mail.proof.field.client"), clientText || "–"]] };
    }
    case "link.created": {
      const details: Array<[string, string]> = [
        [t("mail.proof.field.name"), text(d.name)],
        [t("mail.proof.field.protection"), t(d.protection === "none" ? "mail.proof.protection.none" : "mail.proof.protection.password")],
        [t("mail.proof.field.expires"), typeof d.expireDate === "string" ? formatDate(Date.parse(`${d.expireDate}T12:00:00Z`)) : ""],
      ];
      if (typeof d.size === "number") details.push([t("mail.proof.field.size"), size(d.size)]);
      if (Array.isArray(d.passwordWays) && d.passwordWays.length) details.push([t("mail.proof.field.passwordWays"), d.passwordWays.map(way).join(", ")]);
      if (typeof d.maxAccess === "number") details.push([t("mail.proof.field.maxAccess"), String(d.maxAccess)]);
      return {
        label: t(`mail.proof.event.linkCreated.${text(d.kind) || "file"}`),
        details,
        // The fingerprint names its file: "Angebot.zip SHA-256: 3a7b…".
        mono: typeof d.sha256 === "string" ? [[t("mail.proof.field.sha256", { name: text(d.name) }), d.sha256]] : undefined,
      };
    }
    case "message.sent":
      return {
        label: t("mail.proof.event.mailSent"),
        details: [
          [t("mail.proof.field.subject"), text(d.subject) || "–"],
          [t("mail.proof.field.from"), text(d.from) || "–"],
          [t("mail.proof.field.to"), list(d.to) || "–"],
          ...(list(d.cc) ? [[t("mail.proof.field.cc"), list(d.cc)] as [string, string]] : []),
          ...(list(d.bcc) ? [[t("mail.proof.field.bcc"), list(d.bcc)] as [string, string]] : []),
          [t("mail.proof.field.sentAt"), isoTime(d.sentAt) || "–"],
          ...about,
        ],
        mono: d.messageId ? [[t("mail.proof.field.messageId"), text(d.messageId)]] : undefined,
      };
    case "password.sent":
      return {
        label: t(d.channel === "sms" ? "mail.proof.event.passwordSms" : "mail.proof.event.passwordMail"),
        details: [[t("mail.proof.field.recipient"), text(d.to)], ...about],
      };
    case "sms.status":
      return { label: t(`mail.proof.event.sms.${text(d.status) || "unknown"}`), details: about };
    case "password.fetched":
      return { label: t("mail.proof.event.passwordFetched"), details: [[t("mail.proof.field.accesses"), text(d.count)]] };
    case "message.opened":
      return { label: t("mail.proof.event.messageOpened"), details: [[t("mail.proof.field.accesses"), text(d.count)], ...about] };
    case "download":
      return { label: t("mail.proof.event.download"), details: [[t("mail.proof.field.downloadedAt"), isoTime(d.downloadedAt)], ...about] };
    case "files.received":
      return { label: t("mail.proof.event.filesReceived"), details: [[t("mail.proof.field.files"), list(d.files)], ...about] };
    case "recipient.codeSent":
    case "recipient.codeFailed":
    case "recipient.unknownAddress":
    case "recipient.codeNotSent":
    case "recipient.verified":
    case "link.opened":
      return {
        label: t(`mail.proof.event.${e.type.replace(/\.(\w)/, (_, c: string) => c.toUpperCase())}`),
        details: [
          [t("mail.proof.field.address"), text(d.address)],
          [t("mail.proof.field.ip"), text(d.ip) || "–"],
          ...(d.userAgent ? [[t("mail.proof.field.browser"), text(d.userAgent)] as [string, string]] : []),
          ...about,
        ],
      };
    case "password.deferred":
      return {
        label: t("mail.proof.event.passwordDeferred"),
        details: [[t("mail.proof.field.passwordWays"), (Array.isArray(d.ways) ? d.ways : []).map(way).join(", ")], ...about],
      };
    case "gateway.locked":
      return {
        label: t("mail.proof.event.gatewayLocked"),
        details: [[t("mail.proof.field.attempts"), text(d.attempts)], [t("mail.proof.field.lockedUntil"), isoTime(d.until)], ...about],
      };
    case "proof.reminded":
      return { label: t("mail.proof.event.proofReminded"), details: [] };
    case "proof.sent":
      return { label: t(d.reason === "final" ? "mail.proof.event.proofSentFinal" : "mail.proof.event.proofSentDownload"), details: [] };
    case "link.revoked":
    case "link.expired":
    case "link.gone":
      return { label: t(`mail.proof.event.${e.type.replace(".", "_")}`), details: about };
    default:
      return { label: e.type, details: about };
  }
}

/**
 * The result first, in a frame: sent, confirmed, fetched, when the links end (or ended) and until when the proof is
 * kept – a tick for what happened, a circle for what has not (yet).
 */
function summaryLines(c: ProofCase): PdfLine[] {
  const s = caseSummary(c);
  const kinds = new Set(Object.values(c.links ?? {}).map((l) => l.kind));
  const fetchedLabel = kinds.has("file") ? "downloaded" : kinds.has("request") ? "received" : "opened";
  const row = (mark: PdfLine["mark"], label: string, value: string): PdfLine => ({ text: label, value, valueAt: 130, mark, size: 10, box: true });
  const lastEnd = Object.values(c.links ?? {}).map((l) => l.expireDate).sort().at(-1);
  return [
    { text: t("mail.proof.summary.title"), font: "bold", size: 11, box: true, gapBefore: 14 },
    row(s.sentAt ? "done" : "open", t("mail.proof.summary.sent"), s.sentAt ? time(s.sentAt) : t("mail.proof.summary.notYet")),
    row(s.confirmed ? "done" : "open", t("mail.proof.summary.confirmed"),
      s.confirmed ? t("mail.proof.summary.confirmedValue", { address: s.confirmed.address, time: time(s.confirmed.at) }) : t("mail.proof.summary.notYet")),
    row(s.fetched ? "done" : "open", t(`mail.proof.summary.${fetchedLabel}`),
      s.fetched ? t("mail.proof.summary.fetchedValue", { count: s.fetched, time: time(s.lastFetchedAt ?? 0) }) : t("mail.proof.summary.notYet")),
    s.endsOn
      ? row("open", t("mail.proof.summary.endsOn"), formatDate(Date.parse(`${s.endsOn}T12:00:00Z`)))
      : row("done", t("mail.proof.summary.ended"), lastEnd ? formatDate(Date.parse(`${lastEnd}T12:00:00Z`)) : "–"),
    row("blank", t("mail.proof.summary.keptUntil"), formatDate(s.keptUntil)),
  ];
}

/** The PDF of a case in `language` (the sender's add-in; else LANGUAGE). */
export function proofPdf(c: ProofCase, language: string | undefined, now = new Date(), signed = false): Buffer {
  return inSenderLanguage(language, () => {
    const check = verifyCase(c);
    const lines: PdfLine[] = [
      { text: t("mail.proof.title"), size: 20, font: "bold" },
      { text: t("mail.proof.trackingId", { id: c.id }), size: 12, font: "bold", gapBefore: 4 },
      { text: t("mail.proof.intro", { sender: c.ownerUpn, date: time(c.createdAt) }), size: 9, grey: 0.3, gapBefore: 6 },
      ...summaryLines(c),
      {
        text: check.intact
          ? t("mail.proof.chainIntact", { count: c.events.length })
          : t("mail.proof.chainBroken", { number: (check.brokenAt ?? 0) + 1 }),
        font: "bold",
        size: 10,
        gapBefore: 16,
      },
    ];
    if (c.events.length) {
      lines.push({ text: `${t("mail.proof.field.lastHash")}: ${c.events.at(-1)!.hash}`, font: "mono", size: 7, grey: 0.35 });
    }
    lines.push({ text: t("mail.proof.events"), size: 13, font: "bold", gapBefore: 16 });
    c.events.forEach((e, i) => {
      const { label, details, mono } = describe(e, c);
      lines.push({ text: `${i + 1}. ${time(e.at)} – ${label}`, font: "bold", size: 10, gapBefore: 8, rule: true });
      for (const [name, value] of details) if (value) lines.push({ text: `${name}: ${value}`, size: 9, indent: 14 });
      for (const [name, value] of mono ?? []) lines.push({ text: `${name}: ${value}`, size: 8, font: "mono", indent: 14 });
      lines.push({ text: `${t("mail.proof.field.hash")}: ${e.hash}`, size: 7, font: "mono", grey: 0.45, indent: 14 });
    });
    // Without Vaultwarden nobody counts reading the password (it was in the mail itself).
    lines.push({ text: t(config.vaultwardenEnabled ? "mail.proof.explanation" : "mail.proof.explanationNoVault", { zone: config.timeZone }),
      size: 8, grey: 0.35, gapBefore: 18, rule: true });
    lines.push({ text: t("mail.proof.generated", { date: time(now.getTime()) }), size: 8, grey: 0.35, gapBefore: 4 });
    return pdfDocument(lines, {
      title: t("mail.proof.documentTitle", { id: c.id }),
      author: c.ownerUpn,
      created: now,
      footer: (page, pages) => t("mail.proof.footer", { id: c.id, page, pages }),
      ...(signed ? { signature: { name: "CloudShare", reason: t("mail.proof.documentTitle", { id: c.id }) } } : {}),
    });
  });
}

/**
 * The PDF signed (pdfSign.ts): any change afterwards shows in the reader. If signing fails, the PDF without a
 * signature field rather than none.
 */
export async function signedProofPdf(c: ProofCase, language: string | undefined, log?: { warn: (o: object, m: string) => void }): Promise<Buffer> {
  const now = new Date();
  try {
    return await signPdf(proofPdf(c, language, now, true));
  } catch (err) {
    log?.warn({ err, case: c.id }, "Delivery proof PDF not signed");
    return proofPdf(c, language, now);
  }
}
