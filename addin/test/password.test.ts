import { describe, expect, it } from "vitest";
import { checkPassword, GENERATED_LENGTH, generatePassword, PASSWORD_MIN_LENGTH } from "../src/shared/password";

describe("generatePassword", () => {
  const samples = Array.from({ length: 300 }, () => generatePassword());

  /** The parts of a password: words "Fuvo"/"Ranto", the short syllable "Be"/"Bem" and the number. */
  const parts = (pw: string) => pw.match(/[A-Z][a-z]+|[2-9]{2}/g) ?? [];
  const WORD = /^[A-Z][aeiou][mnrst]?[a-z][aeiou]$/;
  const SHORT = /^[A-Z][aeiou][mnrst]?$/;

  it("is pronounceable: two words, a short syllable and two digits – 12 to 15 characters, no special characters", () => {
    expect(GENERATED_LENGTH).toEqual({ min: 12, max: 15 });
    for (const pw of samples) {
      expect(pw.length).toBeGreaterThanOrEqual(GENERATED_LENGTH.min);
      expect(pw.length).toBeLessThanOrEqual(GENERATED_LENGTH.max);
      expect(pw).toMatch(/^[A-Za-z2-9]+$/);
      const found = parts(pw);
      expect(found.join("")).toBe(pw);
      expect(found).toHaveLength(4);
      expect(found.filter((p) => /^[2-9]{2}$/.test(p))).toHaveLength(1);
      expect(found.filter((p) => WORD.test(p) && p.length >= 4)).toHaveLength(2);
      expect(found.filter((p) => SHORT.test(p))).toHaveLength(1);
    }
  });

  it("varies: order of the parts, lengths and closed syllables differ", () => {
    const digitsAt = new Set(samples.map((pw) => parts(pw).findIndex((p) => /^\d+$/.test(p))));
    expect([...digitsAt].sort()).toEqual([0, 1, 2, 3]); // the number is not always at the end
    expect(new Set(samples.map((pw) => pw.length)).size).toBeGreaterThanOrEqual(3);
    expect(samples.some((pw) => /[A-Z][aeiou][mnrst][a-z][aeiou]/.test(pw))).toBe(true); // "Ranto"
    expect(samples.some((pw) => /^[A-Z][aeiou][a-z][aeiou][A-Z]/.test(pw))).toBe(true); // "FuvoMo…"
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
