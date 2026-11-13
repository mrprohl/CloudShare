import { describe, expect, it } from "vitest";
import { clampInt, isEmailAddress } from "../src/limits.js";

describe("clampInt (link validity, password-link accesses)", () => {
  it("keeps values within the range", () => {
    expect(clampInt(5, 3, 1, 20)).toBe(5);
    expect(clampInt("7", 3, 1, 20)).toBe(7);
    expect(clampInt(4.9, 3, 1, 20)).toBe(4);
  });

  it("limits values outside the range", () => {
    expect(clampInt(500, 3, 1, 20)).toBe(20);
    expect(clampInt(-2, 3, 1, 20)).toBe(1);
  });

  it("uses the default for missing or invalid input", () => {
    expect(clampInt(undefined, 3, 1, 20)).toBe(3);
    expect(clampInt("abc", 3, 1, 20)).toBe(3);
    expect(clampInt(0, 3, 1, 20)).toBe(3);
  });
});

describe("isEmailAddress (Kennwort an andere Adresse)", () => {
  it.each(["max@example.com", "Max.Muster+cs@sub.example.co.uk", "a@b.de"])("accepts %s", (value) => {
    expect(isEmailAddress(value)).toBe(true);
  });

  it.each(["", "max", "max@", "@example.com", "max@example", "max @example.com", "a@b.de, c@d.de", "a@b.de;c@d.de",
    "Max <max@example.com>", `${"x".repeat(250)}@example.com`])("rejects %j", (value) => {
    expect(isEmailAddress(value)).toBe(false);
  });
});
