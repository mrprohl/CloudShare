/**
 * getToken(): Office allows only one getAccessToken call at a time (error 13008 otherwise). Concurrent callers must
 * share one request, valid tokens are reused, and a 13008 from another runtime is retried.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { installOfficeMock } from "./officeMock";

installOfficeMock();
const { getToken, resetTokenCache } = await import("../src/shared/office");
const getAccessToken = (globalThis as unknown as { Office: { auth: { getAccessToken: ReturnType<typeof vi.fn> } } })
  .Office.auth.getAccessToken;

/** Unsigned JWT that expires in `seconds`. */
const jwt = (seconds: number) =>
  ["{}", JSON.stringify({ exp: Math.floor(Date.now() / 1000) + seconds })]
    .map((part) => Buffer.from(part).toString("base64url"))
    .join(".") + ".sig";

beforeEach(() => {
  resetTokenCache();
  getAccessToken.mockReset();
});

describe("getToken", () => {
  it("shares one Office request between concurrent callers", async () => {
    const token = jwt(3600);
    getAccessToken.mockImplementation(() => new Promise((resolve) => setTimeout(() => resolve(token), 20)));
    const results = await Promise.all([getToken(), getToken(), getToken()]);
    expect(results).toEqual([token, token, token]);
    expect(getAccessToken).toHaveBeenCalledTimes(1);
  });

  it("reuses a valid token and fetches a new one shortly before expiry", async () => {
    getAccessToken.mockResolvedValueOnce(jwt(3600)).mockResolvedValueOnce(jwt(60)).mockResolvedValueOnce(jwt(3600));
    await getToken();
    await getToken();
    expect(getAccessToken).toHaveBeenCalledTimes(1);

    resetTokenCache();
    await getToken(); // expires in 60 s – too close, not reused
    await getToken();
    expect(getAccessToken).toHaveBeenCalledTimes(3);
  });

  it("retries when another runtime is signing in at the same moment (13008)", async () => {
    const token = jwt(3600);
    getAccessToken.mockRejectedValueOnce({ code: 13008, name: "Error", message: "in progress" }).mockResolvedValueOnce(token);
    await expect(getToken()).resolves.toBe(token);
    expect(getAccessToken).toHaveBeenCalledTimes(2);
  });

  it("gives up with a readable error after repeated 13008 and on other codes", async () => {
    getAccessToken.mockRejectedValue({ code: 13008, name: "Error", message: "in progress" });
    await expect(getToken()).rejects.toThrow(/Eine Anmeldung läuft bereits.*\(Code 13008\)/);
    expect(getAccessToken).toHaveBeenCalledTimes(4);

    getAccessToken.mockReset();
    getAccessToken.mockRejectedValue({ code: 13007, name: "Error", message: "x" });
    await expect(getToken()).rejects.toThrow(/\(Code 13007\)/);
    expect(getAccessToken).toHaveBeenCalledTimes(1);
  }, 10_000);
});
