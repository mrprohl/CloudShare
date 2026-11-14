import { describe, expect, it } from "vitest";
import { checkPassword, GENERATED_LENGTH, generatePassword, PASSWORD_MIN_LENGTH } from "../src/shared/password";

describe("generatePassword", () => {
  const samples = Array.from({ length: 300 }, () => generatePassword());

  it("is pronounceable: WortWortSi12 – 12 characters, no special characters", () => {
    for (const pw of samples) {
      expect(pw).toHaveLength(GENERATED_LENGTH);
      expect(GENERATED_LENGTH).toBe(12);
      expect(pw).toMatch(/^[A-Z][aeiou][a-z][aeiou][A-Z][aeiou][a-z][aeiou][A-Z][aeiou][2-9]{2}$/);
    }
  });

  it("avoids easily confused and awkward characters", () => {
    for (const pw of samples) expect(pw).not.toMatch(/[0O1lIcCqQxXyY'"`\\\s!#%*+\-=?@_]/);
  });

  it("always passes its own rules", () => {
    for (const pw of samples) expect(checkPassword(pw).ok).toBe(true);
  });

  it("does not repeat itself", () => {
    expect(new Set(samples).size).toBe(samples.length);
  });

  it("uses every consonant, vowel and digit (no modulo bias)", () => {
    const seen = new Set<string>();
    for (let i = 0; i < 400; i++) for (const c of generatePassword()) seen.add(c.toLowerCase());
    expect([...seen].sort().join("")).toBe([..."bdfghjkmnprstvwz" + "aeiou" + "23456789"].sort().join(""));
  });
});

describe("checkPassword (same rules as the server)", () => {
  it("reports each rule separately for the live check", () => {
    expect(checkPassword("kurz")).toEqual({ length: false, classes: false, variety: false, ok: false });
    expect(checkPassword("nurkleinbuchstaben")).toMatchObject({ length: true, classes: false, ok: false });
    expect(checkPassword("Aa1!Aa1!Aa1!")).toMatchObject({ length: true, classes: true, variety: false, ok: false });
    expect(checkPassword("Sonnenblume-2026!")).toEqual({ length: true, classes: true, variety: true, ok: true });
  });

  it("requires the minimum length", () => {
    expect(PASSWORD_MIN_LENGTH).toBe(12);
    expect(checkPassword("Ab1!Cd2#Ef3").length).toBe(false); // 11 characters
    expect(checkPassword("Ab1!Cd2#Ef3$").length).toBe(true); // 12 characters
  });

  it("rejects simple sequences", () => {
    expect(checkPassword("abcdefghijklmnop").variety).toBe(false);
    expect(checkPassword("ponmlkjihgfedcba").variety).toBe(false);
  });
});
