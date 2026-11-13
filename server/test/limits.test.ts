import { describe, expect, it } from "vitest";
import { clampInt } from "../src/limits.js";

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
