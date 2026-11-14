import { HttpError } from "./auth.js";
import { config } from "./config.js";
import type { NextcloudCredentials } from "./credentials.js";

export const NOT_CONNECTED = "NEXTCLOUD_NOT_CONNECTED";
export const QUOTA_EXCEEDED = "NEXTCLOUD_QUOTA_EXCEEDED";

export interface UserInfo {
  id: string;
  /** Free space in bytes; null = unlimited or unknown. */
  free: number | null;
  /** Used space in bytes; null = unknown. */
  used: number | null;
  /** Total space (quota) in bytes; null = unlimited or unknown. */
  total: number | null;
}

const encodePath = (segments: string[]) => segments.map(encodeURIComponent).join("/");

/**
 * Nextcloud may report URLs with an internal scheme/host (e.g. http behind Caddy).
 * Therefore only path + query are kept and appended to the configured public address.
 */
function toPublicUrl(url: string): string {
  const u = new URL(url, config.nextcloudUrl);
  return config.nextcloudUrl + u.pathname + u.search;
}

// ---------------------------------------------------------------------------
// Login Flow v2 – one-time connection per user, yields an app password
// ---------------------------------------------------------------------------

export interface LoginFlow {
  loginUrl: string;
  pollEndpoint: string;
  pollToken: string;
}

export async function startLoginFlow(): Promise<LoginFlow> {
  const res = await fetch(`${config.nextcloudUrl}/index.php/login/v2`, {
    method: "POST",
    headers: { "User-Agent": config.appName, Accept: "application/json" },
  });
  const json = (await res.json().catch(() => null)) as
    | { login?: string; poll?: { token?: string; endpoint?: string } }
    | null;
  if (!res.ok || !json?.login || !json.poll?.token || !json.poll.endpoint) {
    throw new HttpError(502, `Nextcloud: Login-Flow konnte nicht gestartet werden (HTTP ${res.status})`);
  }
  return {
    loginUrl: toPublicUrl(json.login),
    pollEndpoint: toPublicUrl(json.poll.endpoint),
    pollToken: json.poll.token,
  };
}

/** Returns the credentials once the user has granted access, otherwise null. */
export async function pollLoginFlow(flow: LoginFlow): Promise<NextcloudCredentials | null> {
  const res = await fetch(flow.pollEndpoint, {
    method: "POST",
    headers: { "User-Agent": config.appName, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ token: flow.pollToken }),
  });
  if (res.status === 404) return null;
  const json = (await res.json().catch(() => null)) as { loginName?: string; appPassword?: string } | null;
  if (!res.ok || !json?.loginName || !json.appPassword) {
    throw new HttpError(502, `Nextcloud: Login-Flow-Abfrage fehlgeschlagen (HTTP ${res.status})`);
  }
  return { loginName: json.loginName, appPassword: json.appPassword };
}

// ---------------------------------------------------------------------------
// Access on behalf of a connected user (app password, Basic auth)
// ---------------------------------------------------------------------------

export class NextcloudClient {
  private readonly authorization: string;

  /** onUnauthorized: called when Nextcloud rejects the app password (e.g. revoked). */
  constructor(
    creds: NextcloudCredentials,
    private readonly onUnauthorized?: () => Promise<void> | void,
  ) {
    this.authorization = "Basic " + Buffer.from(`${creds.loginName}:${creds.appPassword}`).toString("base64");
  }

  private async request(
    method: string,
    path: string,
    headers: Record<string, string> = {},
    body?: BodyInit | Buffer,
  ) {
    const res = await fetch(config.nextcloudUrl + path, {
      method,
      headers: { Authorization: this.authorization, "User-Agent": config.appName, ...headers },
      // Node's fetch accepts Buffers; the DOM typings of BodyInit just don't know them.
      body: body as BodyInit | undefined,
    });
    if (res.status === 401) {
      await this.onUnauthorized?.();
      throw new HttpError(409, "Die Nextcloud-Verbindung ist nicht mehr gültig. Bitte neu verbinden.", NOT_CONNECTED);
    }
    return res;
  }

  private async expect(res: Response, action: string, ...okStatus: number[]) {
    if (res.ok || okStatus.includes(res.status)) return;
    if (res.status === 507) {
      throw new HttpError(507, "Nicht genügend Speicherplatz in Ihrer Nextcloud.", QUOTA_EXCEEDED);
    }
    throw new HttpError(502, `Nextcloud: ${action} fehlgeschlagen (HTTP ${res.status})`);
  }

  private async ocs(method: string, path: string, body?: URLSearchParams) {
    const res = await this.request(
      method,
      `${path}${path.includes("?") ? "&" : "?"}format=json`,
      {
        "OCS-APIRequest": "true",
        Accept: "application/json",
        ...(body ? { "Content-Type": "application/x-www-form-urlencoded" } : {}),
      },
      body,
    );
    const json = (await res.json().catch(() => null)) as
      | { ocs?: { meta?: { message?: string }; data?: unknown } }
      | null;
    if (!res.ok || !json?.ocs) {
      const message = json?.ocs?.meta?.message || `HTTP ${res.status}`;
      throw new HttpError(res.status >= 500 ? 502 : 400, `Nextcloud: ${message}`);
    }
    return json.ocs.data;
  }

  private davFiles(uid: string, segments: string[]) {
    return `/remote.php/dav/files/${encodeURIComponent(uid)}/${encodePath(segments)}`;
  }

  private davUpload(uid: string, uploadId: string, chunk?: string) {
    const base = `/remote.php/dav/uploads/${encodeURIComponent(uid)}/${encodeURIComponent(uploadId)}`;
    return chunk ? `${base}/${chunk}` : base;
  }

  private destination(uid: string, segments: string[]) {
    return config.nextcloudUrl + this.davFiles(uid, segments);
  }

  /**
   * Whether Nextcloud enforces passwords for public links for this user (Settings → Sharing; capabilities are
   * evaluated per user, so group exceptions count). Then "Ohne Kennwort" cannot work.
   */
  async linkPasswordEnforced(): Promise<boolean> {
    const data = (await this.ocs("GET", "/ocs/v2.php/cloud/capabilities")) as {
      capabilities?: { files_sharing?: { public?: { password?: { enforced?: boolean } } } };
    };
    return data?.capabilities?.files_sharing?.public?.password?.enforced === true;
  }

  /** Internal user ID (may differ from the login name, e.g. with user_oidc) and free quota. */
  async userInfo(): Promise<UserInfo> {
    const data = (await this.ocs("GET", "/ocs/v2.php/cloud/user")) as {
      id?: string;
      quota?: { free?: number; used?: number; total?: number; quota?: number };
    };
    if (!data?.id) throw new HttpError(502, "Nextcloud: Benutzer-ID nicht ermittelbar");
    // Negative values mean "unlimited" or "not computed" in Nextcloud.
    const bytes = (value: unknown) => (typeof value === "number" && value >= 0 ? value : null);
    const free = bytes(data.quota?.free);
    const used = bytes(data.quota?.used);
    // With a quota "total" is the quota itself; without one Nextcloud reports free disk space + used.
    const total = bytes(data.quota?.quota) !== null ? bytes(data.quota?.total) : null;
    return { id: data.id, free, used, total };
  }

  /** Revokes the app password this client is using. */
  async revokeAppPassword() {
    await this.ocs("DELETE", "/ocs/v2.php/core/apppassword");
  }

  /** Deletes a file or folder (it goes to the user's Nextcloud trash); already gone (404) is fine. */
  async deletePath(uid: string, segments: string[]) {
    const res = await this.request("DELETE", this.davFiles(uid, segments));
    await this.expect(res, "Löschen", 404);
  }

  /** Creates the folder chain; existing folders (405) are fine. */
  async ensureFolder(uid: string, segments: string[]) {
    for (let i = 1; i <= segments.length; i++) {
      const res = await this.request("MKCOL", this.davFiles(uid, segments.slice(0, i)));
      await this.expect(res, "Ordner anlegen", 405);
    }
  }

  async startUpload(uid: string, uploadId: string, target: string[]) {
    const res = await this.request("MKCOL", this.davUpload(uid, uploadId), {
      Destination: this.destination(uid, target),
    });
    await this.expect(res, "Upload starten");
  }

  async putChunk(uid: string, uploadId: string, index: number, target: string[], data: Buffer) {
    const res = await this.request(
      "PUT",
      this.davUpload(uid, uploadId, String(index)),
      { Destination: this.destination(uid, target), "Content-Type": "application/octet-stream" },
      data,
    );
    await this.expect(res, `Block ${index} hochladen`);
  }

  async finishUpload(uid: string, uploadId: string, target: string[], totalLength: number) {
    const res = await this.request("MOVE", this.davUpload(uid, uploadId, ".file"), {
      Destination: this.destination(uid, target),
      "OC-Total-Length": String(totalLength),
    });
    await this.expect(res, "Upload abschließen");
  }

  async abortUpload(uid: string, uploadId: string) {
    await this.request("DELETE", this.davUpload(uid, uploadId)).catch(() => undefined);
  }

  /**
   * Public link with password and expiry date (YYYY-MM-DD).
   * permissions: 1 = read only (download link), 4 = create only ("file drop" upload link for a folder).
   */
  async createPublicShare(
    target: string[],
    /** undefined: link without password (only for file links "Ohne Kennwort"). */
    password: string | undefined,
    expireDate: string,
    permissions: 1 | 4 = 1,
    label?: string,
  ): Promise<string> {
    const params = new URLSearchParams({
      path: "/" + target.join("/"),
      shareType: "3",
      permissions: String(permissions),
      expireDate,
    });
    if (password !== undefined) params.set("password", password);
    if (label) params.set("label", label);
    const data = (await this.ocs("POST", "/ocs/v2.php/apps/files_sharing/api/v1/shares", params)) as { url?: string };
    if (!data?.url) throw new HttpError(502, "Nextcloud: Freigabelink fehlt in der Antwort");
    return data.url;
  }

  /** All public link shares owned by the user. */
  async listPublicShares(): Promise<PublicShare[]> {
    const data = (await this.ocs("GET", "/ocs/v2.php/apps/files_sharing/api/v1/shares")) as Array<{
      id: string | number;
      share_type: number;
      path: string;
      url?: string;
      expiration?: string | null;
      stime: number;
      item_source: number;
      item_type: string;
      permissions: number;
      label?: string;
      item_size?: number;
      /** null without password, otherwise "redacted" (older versions: the hash). */
      password?: string | null;
    }>;
    return (data ?? [])
      .filter((s) => Number(s.share_type) === 3)
      .map((s) => ({
        id: String(s.id),
        path: s.path,
        url: s.url ?? "",
        expiration: s.expiration ? s.expiration.slice(0, 10) : null,
        sharedAt: Number(s.stime) * 1000,
        fileId: Number(s.item_source),
        isFolder: s.item_type === "folder",
        permissions: Number(s.permissions),
        label: s.label ?? "",
        size: typeof s.item_size === "number" && s.item_size >= 0 ? s.item_size : null,
        hasPassword: Boolean(s.password),
      }));
  }

  async deleteShare(id: string) {
    await this.ocs("DELETE", `/ocs/v2.php/apps/files_sharing/api/v1/shares/${encodeURIComponent(id)}`);
  }

  /**
   * Downloads of a public link according to Nextcloud's activity app (type "public_links").
   * Returns null if the activity app is not available.
   */
  async publicDownloads(fileId: number): Promise<{ count: number; last: string | null } | null> {
    const res = await this.request(
      "GET",
      `/ocs/v2.php/apps/activity/api/v2/activity/filter?format=json&object_type=files&object_id=${fileId}&limit=100`,
      { "OCS-APIRequest": "true", Accept: "application/json" },
    );
    if (res.status === 304) return { count: 0, last: null }; // no activities at all
    if (!res.ok) return null;
    const json = (await res.json().catch(() => null)) as {
      ocs?: { data?: Array<{ type?: string; subject?: string; datetime?: string }> };
    } | null;
    // Type "public_links" only contains downloads through public links (files_sharing's "Downloads" provider).
    // The subject is localised ("…was downloaded" / "…wurde heruntergeladen"), so it must not be parsed.
    const downloads = (json?.ocs?.data ?? []).filter((a) => a.type === "public_links");
    const last = downloads.map((a) => a.datetime ?? "").sort().pop() || null;
    return { count: downloads.length, last };
  }

  /** Number of files in a folder (e.g. uploads received through an upload link). */
  async countFiles(uid: string, segments: string[]): Promise<number | null> {
    const res = await this.request("PROPFIND", this.davFiles(uid, segments), {
      Depth: "1",
      "Content-Type": "application/xml",
    }, '<?xml version="1.0"?><d:propfind xmlns:d="DAV:"><d:prop><d:resourcetype/></d:prop></d:propfind>');
    if (!res.ok) return null;
    const xml = await res.text();
    const responses = xml.match(/<d:response>/g)?.length ?? 0;
    const folders = xml.match(/<d:collection\s*\/>/g)?.length ?? 0;
    // The folder itself is one response and one collection.
    return Math.max(0, responses - folders);
  }
}

export interface PublicShare {
  id: string;
  /** Path in the user's files, e.g. "/CloudShare/2026-11-13_ab12cd34/report.pdf". */
  path: string;
  url: string;
  /** YYYY-MM-DD or null. */
  expiration: string | null;
  sharedAt: number;
  fileId: number;
  isFolder: boolean;
  permissions: number;
  label: string;
  /** File size in bytes (Nextcloud "item_size"); null if not reported. */
  size: number | null;
  /** Whether the link is password-protected (Nextcloud does not tell random from own passwords). */
  hasPassword: boolean;
}
