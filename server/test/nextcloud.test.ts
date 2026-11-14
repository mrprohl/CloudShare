import { afterEach, describe, expect, it, vi } from "vitest";
import { HttpError } from "../src/auth.js";
import { NextcloudClient, NOT_CONNECTED, QUOTA_EXCEEDED, startLoginFlow } from "../src/nextcloud.js";
import { json, mockFetch, ocs } from "./helpers.js";

const creds = { loginName: "max", appPassword: "app-pass" };

afterEach(() => vi.unstubAllGlobals());

describe("NextcloudClient", () => {
  it("uses Basic auth with the app password and the OCS header", async () => {
    const { call } = mockFetch(ocs({ id: "max-uid", quota: { free: 5_000_000 } }));
    const info = await new NextcloudClient(creds).userInfo();
    expect(info).toEqual({ id: "max-uid", free: 5_000_000, used: null, total: null });
    const headers = call(0).init.headers as Record<string, string>;
    expect(headers.Authorization).toBe(`Basic ${Buffer.from("max:app-pass").toString("base64")}`);
    expect(headers["OCS-APIRequest"]).toBe("true");
    expect(call(0).url).toBe("https://cloud.example.test/ocs/v2.php/cloud/user?format=json");
  });

  it("reports free, used and total space for a user with a quota", async () => {
    mockFetch(ocs({ id: "max-uid", quota: { free: 30, used: 20, total: 50, quota: 50 } }));
    expect(await new NextcloudClient(creds).userInfo()).toEqual({ id: "max-uid", free: 30, used: 20, total: 50 });
  });

  it("has no total without a quota (Nextcloud's total is then just disk space)", async () => {
    mockFetch(ocs({ id: "max-uid", quota: { free: 900, used: 20, total: 920, quota: -3 } }));
    expect(await new NextcloudClient(creds).userInfo()).toEqual({ id: "max-uid", free: 900, used: 20, total: null });
  });

  it("reads the enforced link password from the capabilities", async () => {
    const { call } = mockFetch(
      ocs({ capabilities: { files_sharing: { public: { password: { enforced: true } } } } }),
      ocs({ capabilities: { files_sharing: { public: { password: { enforced: false } } } } }),
      ocs({ capabilities: {} }),
    );
    const client = new NextcloudClient(creds);
    expect(await client.linkPasswordEnforced()).toBe(true);
    expect(await client.linkPasswordEnforced()).toBe(false);
    expect(await client.linkPasswordEnforced()).toBe(false);
    expect(call(0).url).toBe("https://cloud.example.test/ocs/v2.php/cloud/capabilities?format=json");
  });

  it("treats negative quota values as unlimited", async () => {
    mockFetch(ocs({ id: "max-uid", quota: { free: -3 } }));
    expect((await new NextcloudClient(creds).userInfo()).free).toBeNull();
  });

  it("reports a revoked app password and lets the caller forget it", async () => {
    mockFetch(new Response("", { status: 401 }));
    const forget = vi.fn();
    const error = await new NextcloudClient(creds, forget).userInfo().catch((e) => e);
    expect(error).toBeInstanceOf(HttpError);
    expect(error.code).toBe(NOT_CONNECTED);
    expect(forget).toHaveBeenCalledOnce();
  });

  it("maps 507 to a quota error", async () => {
    mockFetch(new Response("", { status: 507 }));
    const error = await new NextcloudClient(creds).putChunk("max-uid", "u1", 1, ["A", "b.bin"], Buffer.from("x")).catch((e) => e);
    expect(error.code).toBe(QUOTA_EXCEEDED);
  });

  it("creates read-only and upload-only public links with password and expiry", async () => {
    const { call } = mockFetch(ocs({ url: "https://cloud.example.test/s/abc" }), ocs({ url: "https://cloud.example.test/s/up" }));
    const client = new NextcloudClient(creds);
    expect(await client.createPublicShare(["CloudShare", "x", "a.pdf"], "pw", "2026-11-27")).toBe(
      "https://cloud.example.test/s/abc",
    );
    await client.createPublicShare(["CloudShare-Dateianforderungen", "y"], "pw", "2026-11-27", 4, "Dateianforderung");

    const first = new URLSearchParams(String(call(0).init.body));
    expect(Object.fromEntries(first)).toEqual({
      path: "/CloudShare/x/a.pdf",
      shareType: "3",
      permissions: "1",
      password: "pw",
      expireDate: "2026-11-27",
    });
    const second = new URLSearchParams(String(call(1).init.body));
    expect(second.get("permissions")).toBe("4");
    expect(second.get("label")).toBe("Dateianforderung");
  });

  it("creates a link without password when none is given (file links 'Ohne Kennwort')", async () => {
    const { call } = mockFetch(ocs({ url: "https://cloud.example.test/s/open" }));
    await new NextcloudClient(creds).createPublicShare(["CloudShare", "x", "a.pdf"], undefined, "2026-11-27");
    const params = new URLSearchParams(String(call(0).init.body));
    expect(params.has("password")).toBe(false);
    expect(params.get("expireDate")).toBe("2026-11-27");
  });

  it("passes Nextcloud's own error message through (e.g. password policy)", async () => {
    mockFetch(ocs(null, 400, "Das Passwort ist zu häufig verwendet"));
    await expect(new NextcloudClient(creds).createPublicShare(["a"], "pw", "2026-11-27")).rejects.toThrow(
      "Nextcloud: Das Passwort ist zu häufig verwendet",
    );
  });

  it("lists only public links", async () => {
    mockFetch(
      ocs([
        { id: 1, share_type: 3, path: "/CloudShare/x/a.pdf", url: "u1", expiration: "2026-11-27 00:00:00", stime: 1_794_000_000, item_source: 11, item_type: "file", permissions: 1, item_size: 4096, password: "redacted" },
        { id: 2, share_type: 0, path: "/privat.txt", stime: 1, item_source: 12, item_type: "file", permissions: 1 },
      ]),
    );
    const shares = await new NextcloudClient(creds).listPublicShares();
    expect(shares).toEqual([
      {
        id: "1",
        path: "/CloudShare/x/a.pdf",
        url: "u1",
        expiration: "2026-11-27",
        sharedAt: 1_794_000_000_000,
        fileId: 11,
        isFolder: false,
        permissions: 1,
        label: "",
        size: 4096,
        hasPassword: true,
      },
    ]);
  });

  it("counts public downloads from the activity app", async () => {
    mockFetch(
      ocs([
        { type: "public_links", subject: "Öffentlich geteilte Datei a.pdf wurde heruntergeladen", datetime: "2026-11-13T09:00:00+00:00" },
        { type: "public_links", subject: "Public shared file a.pdf was downloaded", datetime: "2026-11-13T10:20:00+00:00" },
        { type: "file_changed", subject: "You changed a.pdf", datetime: "2026-11-13T11:00:00+00:00" },
      ]),
    );
    expect(await new NextcloudClient(creds).publicDownloads(11)).toEqual({
      count: 2,
      last: "2026-11-13T10:20:00+00:00",
    });
  });

  it("reports zero downloads on 304 and 'unknown' without activity app", async () => {
    mockFetch(new Response(null, { status: 304 }), new Response("", { status: 404 }));
    const client = new NextcloudClient(creds);
    expect(await client.publicDownloads(11)).toEqual({ count: 0, last: null });
    expect(await client.publicDownloads(11)).toBeNull();
  });

  it("counts received files in an upload folder", async () => {
    const xml =
      `<?xml version="1.0"?><d:multistatus xmlns:d="DAV:">` +
      `<d:response><d:href>/f/</d:href><d:propstat><d:prop><d:resourcetype><d:collection/></d:resourcetype></d:prop></d:propstat></d:response>` +
      `<d:response><d:href>/f/a.pdf</d:href><d:propstat><d:prop><d:resourcetype/></d:prop></d:propstat></d:response>` +
      `<d:response><d:href>/f/b.pdf</d:href><d:propstat><d:prop><d:resourcetype/></d:prop></d:propstat></d:response>` +
      `</d:multistatus>`;
    const { call } = mockFetch(new Response(xml, { status: 207 }));
    expect(await new NextcloudClient(creds).countFiles("max-uid", ["CloudShare-Dateianforderungen", "f"])).toBe(2);
    expect(call(0).init.method).toBe("PROPFIND");
  });
});

describe("startLoginFlow", () => {
  it("rewrites URLs reported with an internal scheme onto the public Nextcloud address", async () => {
    mockFetch(
      json({
        login: "http://nextcloud-web/index.php/login/v2/flow/abc",
        poll: { token: "t", endpoint: "http://nextcloud-web/index.php/login/v2/poll" },
      }),
    );
    expect(await startLoginFlow()).toEqual({
      loginUrl: "https://cloud.example.test/index.php/login/v2/flow/abc",
      pollEndpoint: "https://cloud.example.test/index.php/login/v2/poll",
      pollToken: "t",
    });
  });
});
