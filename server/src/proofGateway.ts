/**
 * "Zustellnachweis" – the recipient confirms their address before a link opens. The link in the mail points here
 * (/r/s/<token>) instead of to Nextcloud:
 *
 *   1. The page asks for the recipient's address. Only an address the mail went to (To, Cc, Bcc) gets a code – sent from
 *      the sender's mailbox with the application permission Mail.Send. The page says the same for every address, so
 *      it does not reveal who the recipients are.
 *   2. The recipient enters the six-digit code (10 minutes, 5 tries). The case records the confirmed address, the
 *      address of the client and its browser, then the browser goes on to the Nextcloud link (still with its password).
 *   3. A signed cookie for this link (until it expires) lets the same browser open it again without a new code; every
 *      opening is recorded.
 *
 * Plain HTML forms, no script; the page comes in the language chosen for the recipients.
 */
import { createHash, randomInt } from "node:crypto";
import { appGraphToken } from "./auth.js";
import { config } from "./config.js";
import { sendMailAsApp } from "./graph.js";
import { inRecipientLanguage, mailLanguage, t } from "./i18n.js";
import { escapeHtml, mailDocument, paragraph } from "./mailLayout.js";
import { addEvent, caseOfToken, type ProofCase, type ProofLink } from "./proof.js";
import { sign, signedBy } from "./storage.js";

const MINUTE = 60_000;
const CODE_TTL = 10 * MINUTE;
const MAX_TRIES = 5;
/** Codes per address and link in 15 minutes, and per link in an hour (no mail flood from the sender's mailbox). */
const CODES_PER_ADDRESS = 3;
const CODES_PER_LINK = 10;

interface PendingCode {
  hash: string;
  expiresAt: number;
  tries: number;
}

/** Codes waiting to be entered (token|address); in memory – a restart only means asking for a new code. */
const pending = new Map<string, PendingCode>();
/** When codes were sent, per token|address and per token – for the limits above. */
const sentTimes = new Map<string, number[]>();

const hashCode = (code: string) => createHash("sha256").update(code).digest("hex");
const normalize = (address: unknown) => (typeof address === "string" ? address.trim().toLowerCase() : "");
const validAddress = (address: string) => /^[^\s@]{1,64}@[^\s@]{1,255}$/.test(address);

function recent(key: string, window: number, now: number): number[] {
  const times = (sentTimes.get(key) ?? []).filter((at) => now - at < window);
  sentTimes.set(key, times);
  return times;
}

/** What the browser sent – for the record (address of the client behind Caddy, its browser). */
export interface Visitor {
  ip: string;
  userAgent: string;
}

export interface GatewayResult {
  status: number;
  html?: string;
  /** Go on to this address (the Nextcloud link) and set this cookie. */
  redirect?: string;
  cookie?: string;
}

/** Content-Security-Policy of the pages: nothing but their own styles; the form may lead on to Nextcloud. */
export const gatewayCsp = () =>
  `default-src 'none'; style-src 'unsafe-inline'; img-src data:; form-action 'self' ${new URL(config.nextcloudUrl).origin}; ` +
  "frame-ancestors 'none'; base-uri 'none'";

const COOKIE = "cloudshare_proof";
const cookiePath = (token: string) => `/r/s/${encodeURIComponent(token)}`;

/** The cookie of a confirmed recipient: token, address and end, signed. */
function confirmationCookie(token: string, address: string, link: ProofLink): string {
  const until = Date.parse(`${link.expireDate}T00:00:00Z`) + 86_400_000;
  const value = `${Buffer.from(address).toString("base64url")}.${until}`;
  const signature = sign(`proof-cookie|${token}|${value}`);
  return `${COOKIE}=${value}.${signature}; Path=${cookiePath(token)}; Max-Age=${Math.max(0, Math.floor((until - Date.now()) / 1000))}; HttpOnly; Secure; SameSite=Lax`;
}

/** The address a valid cookie confirms for this link, if any. */
function confirmedAddress(token: string, cookieHeader: string | undefined): string | undefined {
  const raw = (cookieHeader ?? "").split(/;\s*/).find((c) => c.startsWith(`${COOKIE}=`))?.slice(COOKIE.length + 1);
  const [address, until, signature] = (raw ?? "").split(".");
  if (!address || !until || !signedBy(`proof-cookie|${token}|${address}.${until}`, signature)) return undefined;
  if (Number(until) < Date.now()) return undefined;
  return Buffer.from(address, "base64url").toString();
}

/** The case and link behind a token – only Nextcloud links of a case that are still valid. */
function linkOf(token: string): { c: ProofCase; link: ProofLink } | undefined {
  const c = caseOfToken(token);
  const link = c?.links?.[token];
  if (!c || !link?.url || c.ended?.includes(token)) return undefined;
  if (Date.parse(`${link.expireDate}T00:00:00Z`) + 86_400_000 < Date.now()) return undefined;
  return { c, link };
}

const page = (title: string, body: string) => `<!doctype html>
<html lang="${escapeHtml(mailLanguage())}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex"><title>${escapeHtml(title)}</title>
<style>
:root{color-scheme:light dark;--bg:#f5f5f5;--card:#fff;--fg:#242424;--fg2:#616161;--stroke:#d1d1d1;--brand:#0f6cbd;--brand-hover:#115ea3;--danger:#c50f1f}
@media (prefers-color-scheme:dark){:root{--bg:#1f1f1f;--card:#292929;--fg:#fff;--fg2:#adadad;--stroke:#666;--brand:#479ef5;--brand-hover:#62abf5;--danger:#dc626d}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:14px/1.45 "Segoe UI",-apple-system,Helvetica,Arial,sans-serif}
main{max-width:440px;margin:48px auto;padding:0 16px}.card{background:var(--card);border:1px solid var(--stroke);border-left:4px solid var(--brand);border-radius:8px;padding:24px}
h1{font-size:20px;font-weight:600;margin:0 0 12px}p{margin:0 0 12px}.small{font-size:12px;color:var(--fg2)}.error{color:var(--danger);font-weight:600}
label{display:block;font-weight:600;margin:16px 0 4px}input{width:100%;font:inherit;padding:7px 10px;border:1px solid var(--stroke);border-bottom-color:var(--fg2);border-radius:4px;background:var(--card);color:var(--fg)}
input:focus{outline:2px solid var(--brand);outline-offset:-1px}button{margin-top:16px;font:inherit;font-weight:600;color:#fff;background:var(--brand);border:0;border-radius:4px;padding:8px 20px;cursor:pointer}
button:hover{background:var(--brand-hover)}.code{letter-spacing:.3em;font-size:18px}
</style></head><body><main><div class="card"><h1>${escapeHtml(title)}</h1>${body}</div>
<p class="small" style="margin-top:12px">${escapeHtml(t("mail.proof.gateway.privacy"))}</p></main></body></html>`;

const addressForm = (token: string, error = "") =>
  page(
    t("mail.proof.gateway.title"),
    `<p>${escapeHtml(t("mail.proof.gateway.intro"))}</p>${error ? `<p class="error" role="alert">${escapeHtml(error)}</p>` : ""}
<form method="post" action="${cookiePath(token)}/code"><label for="address">${escapeHtml(t("mail.proof.gateway.address"))}</label>
<input id="address" name="address" type="email" autocomplete="email" required autofocus>
<button type="submit">${escapeHtml(t("mail.proof.gateway.sendCode"))}</button></form>`,
  );

const codeForm = (token: string, address: string, error = "") =>
  page(
    t("mail.proof.gateway.codeTitle"),
    `<p>${escapeHtml(t("mail.proof.gateway.codeSent", { address }))}</p>${error ? `<p class="error" role="alert">${escapeHtml(error)}</p>` : ""}
<form method="post" action="${cookiePath(token)}/verify"><input type="hidden" name="address" value="${escapeHtml(address)}">
<label for="code">${escapeHtml(t("mail.proof.gateway.code"))}</label>
<input id="code" class="code" name="code" inputmode="numeric" autocomplete="one-time-code" pattern="[0-9]{6}" maxlength="6" required autofocus>
<button type="submit">${escapeHtml(t("mail.proof.gateway.open"))}</button></form>
<p class="small"><a href="${cookiePath(token)}" style="color:inherit">${escapeHtml(t("mail.proof.gateway.otherAddress"))}</a></p>`,
  );

const unavailable = () =>
  ({ status: 404, html: page(t("mail.proof.gateway.unavailableTitle"), `<p>${escapeHtml(t("mail.proof.gateway.unavailable"))}</p>`) });

/** GET /r/s/:token – a confirmed browser goes on to the link (recorded), any other gets the address form. */
export function openGateway(token: string, cookieHeader: string | undefined, visitor: Visitor): GatewayResult {
  const found = linkOf(token);
  if (!found) return inRecipientLanguage(undefined, unavailable);
  const { c, link } = found;
  return inRecipientLanguage(c.language, () => {
    if (!c.recipients?.length) {
      return { status: 409, html: page(t("mail.proof.gateway.notSentTitle"), `<p>${escapeHtml(t("mail.proof.gateway.notSent"))}</p>`) };
    }
    const address = confirmedAddress(token, cookieHeader);
    if (address && c.recipients.includes(address)) {
      addEvent(c.id, "link.opened", { links: [token], address, ip: visitor.ip, userAgent: visitor.userAgent.slice(0, 300) });
      return { status: 303, redirect: link.url };
    }
    return { status: 200, html: addressForm(token) };
  });
}

/** POST /r/s/:token/code – sends a code if the address is a recipient; the page reads the same either way. */
export async function requestCode(token: string, rawAddress: unknown, visitor: Visitor, now = Date.now()): Promise<GatewayResult> {
  const found = linkOf(token);
  if (!found) return inRecipientLanguage(undefined, unavailable);
  const { c, link } = found;
  const address = normalize(rawAddress);
  if (!validAddress(address)) {
    return inRecipientLanguage(c.language, () => ({ status: 400, html: addressForm(token, t("mail.proof.gateway.invalidAddress")) }));
  }
  if (!c.recipients?.includes(address)) {
    // Not a recipient: no mail – but the same answer, so the page does not tell who the recipients are.
    return inRecipientLanguage(c.language, () => ({ status: 200, html: codeForm(token, address) }));
  }
  const key = `${token}|${address}`;
  if (recent(key, 15 * MINUTE, now).length >= CODES_PER_ADDRESS || recent(token, 60 * MINUTE, now).length >= CODES_PER_LINK) {
    return inRecipientLanguage(c.language, () => ({ status: 429, html: codeForm(token, address, t("mail.proof.gateway.tooMany")) }));
  }
  const code = String(randomInt(0, 1_000_000)).padStart(6, "0");
  const { subject, html } = inRecipientLanguage(c.language, () => ({
    subject: t("mail.proof.codeMail.subject", { code }),
    html: mailDocument({
      opening: paragraph(t("mail.proof.codeMail.text", { name: link.name })),
      card: {
        accent: "brand",
        title: code,
        details: t("mail.proof.codeMail.details", { minutes: CODE_TTL / MINUTE, id: c.id }),
      },
    }),
  }));
  await sendMailAsApp(await appGraphToken(), c.owner, address, subject, html);
  pending.set(key, { hash: hashCode(code), expiresAt: now + CODE_TTL, tries: 0 });
  sentTimes.get(key)!.push(now);
  sentTimes.get(token)!.push(now);
  addEvent(c.id, "recipient.codeSent", { links: [token], address, ip: visitor.ip });
  return inRecipientLanguage(c.language, () => ({ status: 200, html: codeForm(token, address) }));
}

/** POST /r/s/:token/verify – a right code confirms the address (recorded) and leads on to the link. */
export function verifyCode(token: string, rawAddress: unknown, rawCode: unknown, visitor: Visitor, now = Date.now()): GatewayResult {
  const found = linkOf(token);
  if (!found) return inRecipientLanguage(undefined, unavailable);
  const { c, link } = found;
  const address = normalize(rawAddress);
  const code = typeof rawCode === "string" ? rawCode.replace(/\s/g, "") : "";
  const key = `${token}|${address}`;
  const entry = pending.get(key);
  return inRecipientLanguage(c.language, () => {
    if (!entry || entry.expiresAt < now || entry.tries >= MAX_TRIES) {
      pending.delete(key);
      return { status: 400, html: addressForm(token, t("mail.proof.gateway.codeExpired")) };
    }
    if (hashCode(code) !== entry.hash) {
      entry.tries++;
      addEvent(c.id, "recipient.codeFailed", { links: [token], address, ip: visitor.ip });
      return { status: 400, html: codeForm(token, address, t("mail.proof.gateway.wrongCode")) };
    }
    pending.delete(key);
    addEvent(c.id, "recipient.verified", { links: [token], address, ip: visitor.ip, userAgent: visitor.userAgent.slice(0, 300) });
    return { status: 303, redirect: link.url, cookie: confirmationCookie(token, address, link) };
  });
}

/** Tests only. */
export const resetGateway = () => {
  pending.clear();
  sentTimes.clear();
};
