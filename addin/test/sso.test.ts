/** Error texts for Office SSO failures (plain objects from Office.auth.getAccessToken). */
import { describe, expect, it } from "vitest";
import { errorText, SsoError, toSsoError } from "../src/shared/sso";

describe("errorText", () => {
  it("explains known SSO codes in German and keeps the code", () => {
    expect(errorText({ code: 13007, name: "Error", message: "An unexpected error" })).toMatch(
      /kein Anmeldetoken für CloudShare.*\(Code 13007\)$/,
    );
    expect(errorText({ code: "13001" })).toContain("nicht mit Ihrem Microsoft-Konto angemeldet");
  });

  it("falls back to message and code for unknown codes", () => {
    expect(errorText({ code: 50001, message: "Something else" })).toBe("Something else (Code 50001)");
  });

  it("never shows [object Object]", () => {
    expect(errorText({ foo: "bar" })).toBe('{"foo":"bar"}');
    expect(errorText({})).not.toContain("[object Object]");
  });

  it("passes Error messages and strings through", () => {
    expect(errorText(new Error("Kaputt"))).toBe("Kaputt");
    expect(errorText("Text")).toBe("Text");
  });
});

describe("toSsoError", () => {
  it("returns an Error with the SSO code", () => {
    const err = toSsoError({ code: 13005, message: "invalid grant" });
    expect(err).toBeInstanceOf(SsoError);
    expect(err).toBeInstanceOf(Error);
    expect(err.ssoCode).toBe(13005);
    expect(err.message).toContain("(Code 13005)");
  });
});
