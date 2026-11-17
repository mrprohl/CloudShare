/**
 * CloudShare's public links in Nextcloud: which links are CloudShare's and how one is revoked – by its owner
 * ("Zurückziehen" in "Verlauf") or by the job watcher for the links of a mail that was never sent.
 */
import type { FastifyBaseLogger } from "fastify";
import { HttpError } from "./auth.js";
import { expireOnRevoke } from "./cleanup.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import { endLink } from "./proof.js";
import type { UnsentLinks } from "./jobs.js";
import { NextcloudClient, NOT_CONNECTED, type PublicShare } from "./nextcloud.js";
import { forgetMessage, isMessageToken, messageFor } from "./history.js";
import { forgetPasswordToken } from "./passwordRefs.js";
import { deleteSend } from "./vaultwarden.js";

/** Share token: the last segment of a link URL. */
export const tokenOf = (url: string) => url.replace(/\/+$/, "").split("/").pop() ?? "";

/** A link CloudShare created: below the upload or the request folder. */
export const isCloudShareShare = (s: PublicShare) =>
  s.path.startsWith(`/${config.nextcloudFolder}/`) || s.path.startsWith(`/${config.requestFolder}/`);

/** An upload link ("Anfordern"): a folder others may only upload to. */
export const isRequestShare = (s: PublicShare) => s.permissions === 4 && s.isFolder;

type Log = Pick<FastifyBaseLogger, "info" | "warn">;

/**
 * Ends a link; the file itself stays in the owner's Nextcloud. The password is no longer needed once no link of its
 * step is left – its Sends are deleted right away. A revoked file link counts as expired: the upload folder goes to the
 * trash (CLEANUP_AFTER_DAYS) unless another valid link of `shares` still points to it. Received files of an upload
 * link are never deleted. Returns what happens to the file.
 */
export async function revokeShare(
  nc: NextcloudClient,
  owner: string,
  share: PublicShare,
  shares: PublicShare[],
  log: Log,
): Promise<"deleted" | "scheduled" | "kept"> {
  await nc.deleteShare(share.id);
  endLink(tokenOf(share.url), "revoked");
  for (const sendId of await forgetPasswordToken(tokenOf(share.url), owner)) {
    await deleteSend(sendId).catch((err) => log.warn({ err, owner }, "Send of a revoked link not deleted"));
  }
  if (isRequestShare(share)) return "kept";
  const segments = share.path.split("/").filter(Boolean);
  const folder = `/${segments.slice(0, 2).join("/")}/`;
  const today = new Date().toISOString().slice(0, 10);
  const otherLinkActive = shares.some(
    (s) => s.id !== share.id && s.path.startsWith(folder) && (s.expiration === null || s.expiration >= today),
  );
  return expireOnRevoke(owner, segments, otherLinkActive, async (target) => {
    const { id: uid } = await nc.userInfo();
    await nc.deletePath(uid, target);
  });
}

/**
 * Ends a confidential message ("Zurückziehen" in "Verlauf"): its Send is deleted, so the link stops working at once,
 * and the Sends of its password once nothing of the step is left. False if it is not the owner's (or gone already).
 */
export async function revokeMessage(owner: string, token: string, log: Log): Promise<boolean> {
  const message = messageFor(token, owner);
  if (!message) return false;
  await deleteSend(message.sendId);
  endLink(token, "revoked");
  for (const sendId of await forgetPasswordToken(token, owner)) {
    await deleteSend(sendId).catch((err) => log.warn({ err, owner }, "Send of a revoked message's password not deleted"));
  }
  forgetMessage(token);
  return true;
}

/**
 * The links of a mail that was never sent – its password job expired without "Send" ever being pressed: they are
 * revoked and the step's Sends deleted, so no link that never left the draft stays reachable. A link that was already
 * downloaded stays (it reached someone – the mail may have gone out from an Outlook without CloudShare, or the link
 * was copied); the owner decides about it in "Verlauf". Throws if Nextcloud cannot be reached (the watcher
 * tries again later); a user who is no longer connected is skipped.
 */
export async function revokeUnsentLinks(unsent: UnsentLinks, log: Log): Promise<void> {
  const creds = getCredentials(unsent.owner);
  if (!creds) return;
  const nc = new NextcloudClient(creds, () => removeCredentials(unsent.owner));
  try {
    let shares = await nc.listPublicShares();
    const tokens = new Set(unsent.tokens);
    let revoked = 0;
    let kept = 0;
    for (const share of shares.filter((s) => tokens.has(tokenOf(s.url)) && isCloudShareShare(s))) {
      const downloads = isRequestShare(share) ? null : await nc.publicDownloads(share.fileId).catch(() => null);
      if (downloads?.count) {
        kept++;
        continue;
      }
      await revokeShare(nc, unsent.owner, share, shares, log);
      shares = shares.filter((s) => s.id !== share.id);
      revoked++;
    }
    // Confidential messages of that mail: never sent, so nobody has their link.
    for (const token of unsent.tokens.filter(isMessageToken)) {
      if (await revokeMessage(unsent.owner, token, log).catch(() => false)) revoked++;
    }
    for (const sendId of unsent.sendIds) {
      await deleteSend(sendId).catch((err) => log.warn({ err, user: unsent.ownerUpn }, "Send of a mail never sent not deleted"));
    }
    log.info({ user: unsent.ownerUpn, links: revoked, keptDownloaded: kept }, "Links of a mail that was never sent revoked");
  } catch (err) {
    if (err instanceof HttpError && err.code === NOT_CONNECTED) return;
    throw err;
  }
}
