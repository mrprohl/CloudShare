import { afterEach, describe, expect, it, vi } from "vitest";
import { fetchProfile, findSentMessage, JOB_PROPERTY, sendMail } from "../src/graph.js";
import { json, mockFetch } from "./helpers.js";

const CUSTOM_PROPS_ID =
  "String {00020329-0000-0000-C000-000000000046} Name cecp-8f3c2a71-5d4e-4b9a-9e2f-6c1d7a0b3e54";

const sentItem = (jobId: string | null, extra: Record<string, unknown> = {}) => ({
  subject: "Unterlagen",
  from: { emailAddress: { address: "info@example.test" } },
  toRecipients: [{ emailAddress: { address: "erika@example.test" } }],
  ccRecipients: [],
  bccRecipients: [{ emailAddress: { address: "archiv@example.test" } }],
  singleValueExtendedProperties: jobId
    ? [{ id: CUSTOM_PROPS_ID, value: JSON.stringify({ [JOB_PROPERTY]: jobId, other: 1 }) }]
    : undefined,
  ...extra,
});

afterEach(() => vi.unstubAllGlobals());

describe("findSentMessage", () => {
  it("finds the mail tagged with the job ID and returns its real sender and recipients", async () => {
    const { call } = mockFetch(json({ value: [sentItem(null), sentItem("job-2"), sentItem("job-1")] }));
    const sent = await findSentMessage("token", "job-1", Date.UTC(2026, 10, 13));
    expect(sent).toEqual({
      subject: "Unterlagen",
      from: "info@example.test",
      recipients: { to: ["erika@example.test"], cc: [], bcc: ["archiv@example.test"] },
    });

    const { url, init } = call(0);
    expect(url).toContain("https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?");
    expect(decodeURIComponent(url)).toContain("sentDateTime ge 2026-11-13T00:00:00.000Z");
    expect(decodeURIComponent(url)).toContain(`id eq '${CUSTOM_PROPS_ID}'`);
    expect((init.headers as Record<string, string>).Authorization).toBe("Bearer token");
  });

  it("returns null when the mail has not been sent yet", async () => {
    mockFetch(json({ value: [sentItem("job-2")] }));
    expect(await findSentMessage("token", "job-1", 0)).toBeNull();
  });

  it("searches a shared mailbox under /users/{address}", async () => {
    const { call } = mockFetch(json({ value: [] }));
    await findSentMessage("token", "job-1", 0, "info@example.test");
    expect(call(0).url).toContain("/v1.0/users/info%40example.test/mailFolders/sentitems/messages");
  });

  it("ignores broken custom property values", async () => {
    mockFetch(json({ value: [sentItem(null, { singleValueExtendedProperties: [{ id: "x", value: "{kaputt" }] })] }));
    expect(await findSentMessage("token", "job-1", 0)).toBeNull();
  });

  it("throws on Graph errors so the watcher retries", async () => {
    mockFetch(json({ error: { code: "ErrorAccessDenied" } }, 403));
    await expect(findSentMessage("token", "job-1", 0, "fremd@example.test")).rejects.toThrow(/HTTP 403/);
  });
});

describe("sendMail", () => {
  it("sends as the given mailbox when `from` is set", async () => {
    const { call } = mockFetch(new Response(null, { status: 202 }));
    await sendMail("token", { to: ["a@example.test"], cc: [], bcc: [] }, "Kennwort für: X", "<p>x</p>", "info@example.test");
    const body = JSON.parse(String(call(0).init.body));
    expect(body.message.from).toEqual({ emailAddress: { address: "info@example.test" } });
    expect(body.message.toRecipients).toEqual([{ emailAddress: { address: "a@example.test" } }]);
    expect(body.saveToSentItems).toBe(true);
  });

  it("omits `from` for the user's own mailbox", async () => {
    const { call } = mockFetch(new Response(null, { status: 202 }));
    await sendMail("token", { to: ["a@example.test"], cc: [], bcc: [] }, "s", "h");
    expect(JSON.parse(String(call(0).init.body)).message.from).toBeUndefined();
  });

  it("fails when Graph does not accept the mail", async () => {
    mockFetch(json({ error: "nope" }, 400));
    await expect(sendMail("token", { to: [], cc: [], bcc: [] }, "s", "h")).rejects.toThrow(/HTTP 400/);
  });
});

describe("fetchProfile", () => {
  it("maps the Graph user to the footer placeholders", async () => {
    mockFetch(
      json({
        displayName: "Max Mustermann",
        givenName: "Max",
        surname: "Mustermann",
        jobTitle: "Projektleitung",
        department: null,
        companyName: "Muster GmbH",
        mail: null,
        userPrincipalName: "max@example.test",
        businessPhones: ["+49 30 1234567"],
        mobilePhone: null,
        officeLocation: "Berlin",
      }),
    );
    expect(await fetchProfile("token")).toEqual({
      displayName: "Max Mustermann",
      givenName: "Max",
      surname: "Mustermann",
      jobTitle: "Projektleitung",
      department: "",
      company: "Muster GmbH",
      email: "max@example.test", // falls back to the UPN without mail attribute
      phone: "+49 30 1234567",
      mobile: "",
      office: "Berlin",
    });
  });
});
