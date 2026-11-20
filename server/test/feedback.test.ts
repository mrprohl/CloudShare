import { describe, expect, it } from "vitest";
import { feedbackMail, FEEDBACK_MAX_LENGTH, parseFeedback } from "../src/feedback.js";

describe("feedback", () => {
  it("takes a problem or a suggestion with its text; only the known details, cut to length", () => {
    expect(parseFeedback({ kind: "bug", text: "  Der Link fehlt.  " })).toEqual({ kind: "bug", text: "Der Link fehlt." });
    const parsed = parseFeedback({
      kind: "idea", text: "Mehr Farben",
      diagnostics: { outlook: "Outlook 16.0 (PC)", browser: "x".repeat(900), mailbox: "anna@example.test", errors: 42 },
    });
    expect(parsed!.diagnostics).toEqual({ outlook: "Outlook 16.0 (PC)", browser: "x".repeat(500) });
  });

  it("refuses an empty, a too long or an unknown kind of message", () => {
    expect(parseFeedback({ kind: "bug", text: "   " })).toBeNull();
    expect(parseFeedback({ kind: "bug", text: "x".repeat(FEEDBACK_MAX_LENGTH + 1) })).toBeNull();
    expect(parseFeedback({ kind: "spam", text: "Hallo" })).toBeNull();
    expect(parseFeedback(null)).toBeNull();
  });

  it("the mail: who sent it, the text escaped with its lines, the details, and the note before anything public", () => {
    const { subject, html } = feedbackMail(
      { kind: "bug", text: "Schritt 1\nSchritt 2\n\n<b>fett</b>", diagnostics: { outlook: "Outlook 16.0 (PC)" } },
      { name: "Anna Muster", address: "anna@example.test" },
    );
    expect(subject).toBe("CloudShare: Fehlermeldung von Anna Muster");
    expect(html).toContain("Anna Muster (anna@example.test) hat über CloudShare einen Fehler gemeldet.");
    expect(html).toContain("Schritt 1<br>Schritt 2");
    expect(html).toContain("&lt;b&gt;fett&lt;/b&gt;");
    expect(html).not.toContain("<b>fett</b>");
    expect(html).toContain("Outlook 16.0 (PC)");
    expect(html).toContain("Version des Add-ins");
    expect(html).toContain("GitHub-Issue");
    expect(html).toContain("#c50f1f"); // red accent for a problem
  });

  it("without details it says so", () => {
    const { subject, html } = feedbackMail({ kind: "idea", text: "Danke!" }, { name: "Anna Muster", address: "anna@example.test" });
    expect(subject).toBe("CloudShare: Anregung von Anna Muster");
    expect(html).toContain("Ohne technische Angaben gesendet.");
  });
});
