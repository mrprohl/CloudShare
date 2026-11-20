import { describe, expect, it } from "vitest";
import { afterLastCard, withCardAfterLast } from "../src/taskpane/cardPlace";
import { filesCard, requestCard } from "../src/taskpane/mailCard";

const card = (token: string) =>
  filesCard([{ name: `${token}.pdf`, size: 1024, url: `https://cloud.example.com/s/${token}` }], "01.12.2026", () => "1 KB");

describe("a further card right after the last one", () => {
  it("goes after the frame of the card with the last link, not inside it", () => {
    const body = `<html><body><p>Guten Tag,</p>${card("tok1")}<p><br></p><p>Mit freundlichen Grüßen</p></body></html>`;
    const next = requestCard("https://cloud.example.com/s/tok2", "01.12.2026");
    const placed = withCardAfterLast(body, ["tok1"], next)!;
    expect(placed).toContain(`${card("tok1")}${next}<p><br></p><p>Mit freundlichen Grüßen</p>`);
  });

  it("takes the card that comes last in the mail, also inside a body laid out in a table", () => {
    const body = `<table width="100%"><tr><td>${card("tok1")}<p>Text</p>${card("tok2")}<p>Signatur</p></td></tr></table>`;
    const at = afterLastCard(body, ["tok2", "tok1"])!;
    expect(body.slice(0, at).endsWith(card("tok2"))).toBe(true);
    expect(body.slice(at)).toBe("<p>Signatur</p></td></tr></table>");
  });

  it("nothing to go after: no recorded link, link not in the mail, or no card around it (plain text)", () => {
    expect(afterLastCard(card("tok1"), [])).toBeNull();
    expect(afterLastCard(card("tok1"), ["gone"])).toBeNull();
    expect(afterLastCard("<p>https://cloud.example.com/s/tok1</p><table><tr><td>x</td></tr></table>", ["tok1"])).toBeNull();
  });
});
