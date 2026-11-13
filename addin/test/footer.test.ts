import { describe, expect, it } from "vitest";
import { EMPTY_PROFILE, fillFooter, hasPlaceholders } from "../src/taskpane/footer";

const profile = {
  ...EMPTY_PROFILE,
  displayName: "Max Mustermann",
  jobTitle: "Projektleitung",
  company: "Muster GmbH",
  email: "max@example.test",
  phone: "+49 30 1234567",
};
const TEMPLATE = "{displayName} · {jobTitle}\\n{company} · Tel. {phone} · {email}";

describe("fillFooter", () => {
  it("fills placeholders and turns \\n into line breaks", () => {
    expect(fillFooter(TEMPLATE, profile)).toBe(
      "Max Mustermann · Projektleitung\nMuster GmbH · Tel. +49 30 1234567 · max@example.test",
    );
  });

  it("drops a segment with its fixed text when its placeholders are empty", () => {
    expect(fillFooter(TEMPLATE, { ...profile, phone: "" })).toBe(
      "Max Mustermann · Projektleitung\nMuster GmbH · max@example.test",
    );
  });

  it("drops lines that end up empty", () => {
    expect(fillFooter("{department}\\n{displayName}", profile)).toBe("Max Mustermann");
  });

  it("keeps fixed text without placeholders", () => {
    expect(fillFooter("Vertraulich · {department}", profile)).toBe("Vertraulich");
  });

  it("treats unknown placeholders as empty", () => {
    expect(fillFooter("{displayName} · {gibtsNicht}", profile)).toBe("Max Mustermann");
  });

  it("accepts real line breaks too", () => {
    expect(fillFooter("{displayName}\r\n{company}", profile)).toBe("Max Mustermann\nMuster GmbH");
  });

  it("works with a template without placeholders", () => {
    expect(hasPlaceholders("Muster GmbH")).toBe(false);
    expect(hasPlaceholders(TEMPLATE)).toBe(true);
    expect(fillFooter("Muster GmbH", EMPTY_PROFILE)).toBe("Muster GmbH");
  });
});
