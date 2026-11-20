import { describe, expect, it } from "vitest";
import { listContacts, matchContacts, MAX_SUGGESTIONS } from "../src/contacts.js";
import type { Contact } from "../src/graph.js";

const contacts: Contact[] = [
  { name: "Jonas Müller-Berg", emails: ["j.mueller@example.org", "jonas@example.net"], phones: [{ number: "+49 (0) 171 1234567", type: "mobile" }] },
  { name: "Anna Hamüller", emails: ["anna@example.test"], phones: [{ number: "0251 998877", type: "business" }] },
  { name: "Erika Beispiel", emails: ["erika@example.test"], phones: [] },
];

describe("contact suggestions", () => {
  it("finds a name with or without umlaut, a word start before a match inside", () => {
    expect(matchContacts(contacts, "mül", "mail").map((s) => s.value)).toEqual(["j.mueller@example.org", "jonas@example.net", "anna@example.test"]);
    expect(matchContacts(contacts, "mue", "mail")[0]).toEqual({ name: "Jonas Müller-Berg", value: "j.mueller@example.org" });
    expect(matchContacts(contacts, "MULLER", "mail").map((s) => s.name)[0]).toBe("Jonas Müller-Berg");
  });

  it("finds an address and offers every address of a contact once", () => {
    expect(matchContacts(contacts, "erika@", "mail")).toEqual([{ name: "Erika Beispiel", value: "erika@example.test" }]);
    expect(matchContacts(contacts, "jonas", "mail")).toHaveLength(2);
  });

  it("SMS: the digits of the number, with or without country code; contacts without number are left out", () => {
    expect(matchContacts(contacts, "0171", "sms")).toEqual([{ name: "Jonas Müller-Berg", value: "+49 (0) 171 1234567", type: "mobile" }]);
    expect(matchContacts(contacts, "+49 251", "sms").map((s) => s.name)).toEqual(["Anna Hamüller"]);
    expect(matchContacts(contacts, "erika", "sms")).toEqual([]);
  });

  it("needs two characters and stops at the limit", () => {
    expect(matchContacts(contacts, "j", "mail")).toEqual([]);
    const many: Contact[] = Array.from({ length: 20 }, (_, i) => ({ name: `Kontakt ${i}`, emails: [`k${i}@example.test`], phones: [] }));
    expect(matchContacts(many, "kontakt", "mail")).toHaveLength(MAX_SUGGESTIONS);
  });
});

describe("the address book page", () => {
  it("lists every address A to Z, or every number for SMS, one entry each", () => {
    expect(listContacts(contacts, "mail").map((e) => e.value)).toEqual([
      "anna@example.test", "erika@example.test", "j.mueller@example.org", "jonas@example.net",
    ]);
    expect(listContacts(contacts, "sms")).toEqual([
      { name: "Anna Hamüller", value: "0251 998877", type: "business" },
      { name: "Jonas Müller-Berg", value: "+49 (0) 171 1234567", type: "mobile" },
    ]);
  });
});
