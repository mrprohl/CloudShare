// CI only – stand-ins for the external services in the load test (ci/stack.sh load), no dependencies:
//   Microsoft identity platform: GET /jwks (signing key), GET /token?oid=… (signed access token for the add-in)
//   Nextcloud: Login Flow v2, OCS user/shares/app password, WebDAV chunked upload v2 (chunks are counted, not stored)
//   GET /stats – what arrived, so the load test can verify that every byte was uploaded
import { generateKeyPairSync, randomUUID, sign } from "node:crypto";
import { createServer } from "node:http";

const TENANT = process.env.ENTRA_TENANT_ID;
const CLIENT_ID = process.env.ENTRA_CLIENT_ID;
const SELF = "http://mocks:8080";

const { privateKey, publicKey } = generateKeyPairSync("rsa", { modulusLength: 2048 });
const jwk = { ...publicKey.export({ format: "jwk" }), kid: "ci-load", alg: "RS256", use: "sig" };

const b64url = (value) => Buffer.from(typeof value === "string" ? value : JSON.stringify(value)).toString("base64url");

/** Access token as Office SSO would deliver it for this add-in (v2 issuer, audience = client ID). */
function accessToken(oid) {
  const now = Math.floor(Date.now() / 1000);
  const head = b64url({ alg: "RS256", typ: "JWT", kid: jwk.kid });
  const body = b64url({
    aud: CLIENT_ID,
    iss: `https://login.microsoftonline.com/${TENANT}/v2.0`,
    iat: now,
    nbf: now,
    exp: now + 3600,
    scp: "access_as_user",
    oid,
    tid: TENANT,
    preferred_username: `${oid}@example.com`,
  });
  return `${head}.${body}.${sign("RSA-SHA256", Buffer.from(`${head}.${body}`), privateKey).toString("base64url")}`;
}

const stats = { chunks: 0, chunkBytes: 0, uploadsFinished: 0, shares: 0, connects: 0 };

const json = (res, status, body) => {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(body));
};
const ocs = (res, data) => json(res, 200, { ocs: { meta: { status: "ok", statuscode: 200, message: "OK" }, data } });

/** Login name from the Basic auth header (app password of a connected user). */
const loginName = (req) =>
  Buffer.from((req.headers.authorization ?? "").replace(/^Basic /, ""), "base64").toString().split(":")[0];

createServer((req, res) => {
  const url = new URL(req.url, SELF);
  const route = `${req.method} ${url.pathname}`;

  // Request bodies: chunks are only counted; everything else is read completely before answering.
  let bytes = 0;
  let body = "";
  req.on("data", (data) => {
    bytes += data.length;
    if (!url.pathname.startsWith("/remote.php/dav/uploads/")) body += data;
  });
  req.on("end", () => {
    if (route === "GET /jwks") return json(res, 200, { keys: [jwk] });
    if (route === "GET /token") return res.end(accessToken(url.searchParams.get("oid") ?? "load-user"));
    if (route === "GET /stats") return json(res, 200, stats);

    if (route === "POST /index.php/login/v2") {
      const token = randomUUID();
      return json(res, 200, {
        login: `${SELF}/login/v2/flow/${token}`,
        poll: { token, endpoint: `${SELF}/index.php/login/v2/poll` },
      });
    }
    if (route === "POST /index.php/login/v2/poll") {
      // Access is granted immediately – as if the user had already clicked "Allow access".
      const token = new URLSearchParams(body).get("token");
      stats.connects++;
      return json(res, 200, { server: SELF, loginName: `ci-${token.slice(0, 8)}`, appPassword: `app-${token}` });
    }

    if (route === "GET /ocs/v2.php/cloud/user") return ocs(res, { id: loginName(req), quota: { free: 50 * 1024 ** 3 } });
    if (route === "POST /ocs/v2.php/apps/files_sharing/api/v1/shares") {
      stats.shares++;
      return ocs(res, { url: `${SELF}/s/${randomUUID().slice(0, 12)}` });
    }
    if (route === "DELETE /ocs/v2.php/core/apppassword") return ocs(res, []);

    if (req.method === "MKCOL") return res.writeHead(201).end();
    if (req.method === "PUT" && url.pathname.startsWith("/remote.php/dav/uploads/")) {
      stats.chunks++;
      stats.chunkBytes += bytes;
      return res.writeHead(201).end();
    }
    if (req.method === "MOVE") {
      stats.uploadsFinished++;
      return res.writeHead(201).end();
    }
    if (req.method === "DELETE" && url.pathname.startsWith("/remote.php/dav/uploads/")) return res.writeHead(204).end();

    json(res, 404, { error: `mock: ${route} not supported` });
  });
}).listen(8080, "0.0.0.0");
