/**
 * CloudShare's public links in Nextcloud: which links are CloudShare's and how one is revoked – by its owner
 * ("Zurückziehen" in "Meine Freigaben") or by the job watcher for the links of a mail that was never sent.
 */
import type { FastifyBaseLogger } from "fastify";
import { HttpError } from "./auth.js";
import { expireOnRevoke } from "./cleanup.js";
import { config } from "./config.js";
import { getCredentials, removeCredentials } from "./credentials.js";
import type { UnsentLinks } from "./jobs.js";
import { NextcloudClient, NOT_CONNECTED, type PublicShare } from "./nextcloud.js";
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
 * The links of a mail that was never sent – its password job expired without "Send" ever being pressed: they are
 * revoked and the step's Sends deleted, so no link that never left the draft stays reachable. Throws if Nextcloud
 * cannot be reached (the watcher tries again later); a user who is no longer connected is skipped.
 */
export async function revokeUnsentLinks(unsent: UnsentLinks, log: Log): Promise<void> {
  const creds = getCredentials(unsent.owner);
  if (!creds) return;
  const nc = new NextcloudClient(creds, () => removeCredentials(unsent.owner));
  try {
    let shares = await nc.listPublicShares();
    const tokens = new Set(unsent.tokens);
    let revoked = 0;
    for (const share of shares.filter((s) => tokens.has(tokenOf(s.url)) && isCloudShareShare(s))) {
      await revokeShare(nc, unsent.owner, share, shares, log);
      shares = shares.filter((s) => s.id !== share.id);
      revoked++;
    }
    for (const sendId of unsent.sendIds) {
      await deleteSend(sendId).catch((err) => log.warn({ err, user: unsent.ownerUpn }, "Send of a mail never sent not deleted"));
    }
    log.info({ user: unsent.ownerUpn, links: revoked }, "Links of a mail that was never sent revoked");
  } catch (err) {
    if (err instanceof HttpError && err.code === NOT_CONNECTED) return;
    throw err;
  }
}
