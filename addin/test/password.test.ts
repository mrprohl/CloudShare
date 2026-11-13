import { describe, expect, it } from "vitest";
import { checkPassword, GENERATED_LENGTH, generatePassword, PASSWORD_MIN_LENGTH } from "../src/shared/password";

describe("generatePassword", () => {
  const samples = Array.from({ length: 300 }, () => generatePassword());

  it("has the configured length and contains every character class", () => {
    for (const pw of samples) {
      expect(pw).toHaveLength(GENERATED_LENGTH);
      expect(pw).toMatch(/[A-Z]/);
      expect(pw).toMatch(/[a-z]/);
      expect(pw).toMatch(/[0-9]/);
      expect(pw).toMatch(/[!#%*+\-=?@_]/);
    }
  });

  it("avoids easily confused and awkward characters", () => {
    for (const pw of samples) expect(pw).not.toMatch(/[0O1lI'"`\\\s]/);
  });

  it("always passes its own rules", () => {
    for (const pw of samples) expect(checkPassword(pw).ok).toBe(true);
  });

  it("does not repeat itself", () => {
    expect(new Set(samples).size).toBe(samples.length);
  });

  it("uses all symbols of the alphabet roughly evenly (no modulo bias)", () => {
    const counts = new Map<string, number>();
    for (let i = 0; i < 400; i++) for (const c of generatePassword()) counts.set(c, (counts.get(c) ?? 0) + 1);
    expect(counts.size).toBe(67); // 24 upper + 25 lower + 8 digits + 10 symbols
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
