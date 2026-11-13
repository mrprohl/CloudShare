import { describe, expect, it } from "vitest";
import { passwordProblem } from "../src/password.js";

describe("passwordProblem", () => {
  it("accepts strong passwords", () => {
    expect(passwordProblem("kT7#pQ2m!xV9wR4e_Hn3")).toBeNull();
    expect(passwordProblem("Sonnenblume-2026!")).toBeNull();
    expect(passwordProblem("Grüße aus Köln 42")).toBeNull(); // umlauts and spaces count
  });

  it("rejects passwords shorter than 12 characters", () => {
    expect(passwordProblem("Ab1!xyz")).toMatch(/mindestens 12 Zeichen/);
  });

  it("requires at least three character classes", () => {
    expect(passwordProblem("nurkleinbuchstaben")).toMatch(/drei der folgenden Zeichenarten/);
    expect(passwordProblem("kleinundzahlen123")).toMatch(/drei der folgenden Zeichenarten/);
    expect(passwordProblem("Klein und Gross mit Leer")).toBeNull(); // upper, lower, space
  });

  it("rejects repetitions and simple sequences", () => {
    expect(passwordProblem("Aa1!Aa1!Aa1!")).toMatch(/zu einfach/); // fewer than 6 different characters
    expect(passwordProblem("0123456789:;")).toMatch(/zu einfach|Zeichenarten/);
    expect(passwordProblem("abcdefghijklmnop".toUpperCase())).not.toBeNull();
  });
});
