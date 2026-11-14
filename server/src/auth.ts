import { ConfidentialClientApplication } from "@azure/msal-node";
import type { FastifyRequest } from "fastify";
import { createRemoteJWKSet, jwtVerify } from "jose";
import { config } from "./config.js";
import { t } from "./i18n.js";

/** Error returned to the add-in; the message is shown to the user (German). */
export class HttpError extends Error {
  constructor(
    public readonly status: number,
    message: string,
    /** Machine-readable error code for the add-in, e.g. NEXTCLOUD_NOT_CONNECTED. */
    public readonly code?: string,
  ) {
    super(message);
  }
}

export interface User {
  oid: string;
  upn: string;
  /** The add-in's original SSO token, basis for On-Behalf-Of. */
  token: string;
}

const jwks = createRemoteJWKSet(new URL(config.jwksUrl));
const appIdUri = `api://${new URL(config.publicBaseUrl).host}/${config.clientId}`;

export async function authenticate(req: FastifyRequest): Promise<User> {
  const header = req.headers.authorization;
  if (!header?.startsWith("Bearer ")) throw new HttpError(401, t("server.loginRequired"));
  const token = header.slice("Bearer ".length);

  let payload;
  try {
    ({ payload } = await jwtVerify(token, jwks, {
      audience: [config.clientId, appIdUri],
      issuer: [
        `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
        `https://sts.windows.net/${config.tenantId}/`,
      ],
    }));
  } catch {
    throw new HttpError(401, t("server.invalidToken"));
  }

  const scopes = String(payload.scp ?? "").split(" ");
  if (!scopes.includes("access_as_user")) throw new HttpError(403, t("server.scopeMissing"));

  return {
    oid: String(payload.oid),
    upn: String(payload.preferred_username ?? payload.upn ?? ""),
    token,
  };
}

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: config.clientId,
    clientSecret: config.clientSecret,
    authority: `https://login.microsoftonline.com/${config.tenantId}`,
  },
});

/**
 * Token of CloudShare itself (client credentials) for Microsoft Graph – only for the download confirmation, which
 * may happen days after the sender's last token expired. Needs the application permission Mail.Send.
 */
export async function appGraphToken(): Promise<string> {
  try {
    const result = await cca.acquireTokenByClientCredential({ scopes: ["https://graph.microsoft.com/.default"] });
    if (!result?.accessToken) throw new Error("empty response");
    return result.accessToken;
  } catch (err) {
    throw new HttpError(403, t("server.appTokenUnavailable", { error: (err as Error).message }));
  }
}

/** Exchanges the user's SSO token for a token for another API (Microsoft Graph). */
export async function onBehalfOf(user: User, scopes: string[]): Promise<{ token: string; expiresAt: number }> {
  try {
    const result = await cca.acquireTokenOnBehalfOf({ oboAssertion: user.token, scopes });
    if (!result?.accessToken) throw new Error("empty response");
    return {
      token: result.accessToken,
      // Without expiry information assume the usual minimum lifetime of one hour.
      expiresAt: result.expiresOn ? result.expiresOn.getTime() : Date.now() + 60 * 60_000,
    };
  } catch (err) {
    throw new HttpError(
      403,
      t("server.tokenUnavailable", { scopes: scopes.join(" "), error: (err as Error).message }),
    );
  }
}
