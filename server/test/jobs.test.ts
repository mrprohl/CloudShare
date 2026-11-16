import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

// The watcher talks to Graph only through these two functions – replace them with test doubles.
vi.mock("../src/graph.js", () => ({
  findSentMessage: vi.fn(),
  sendMail: vi.fn(),
}));
// SMS channel: the password comes from the Vaultwarden Send, the SMS goes to seven.io – both replaced.
vi.mock("../src/vaultwarden.js", () => ({
  readSendPassword: vi.fn(async () => "KateGeteFi45"),
  deleteSend: vi.fn(async () => undefined),
}));
vi.mock("../src/sms.js", async (importOriginal) => ({
  ...(await importOriginal<typeof import("../src/sms.js")>()),
  sendSms: vi.fn(async () => "77000001"),
}));

import type { UnsentLinks } from "../src/jobs.js";
const graph = await import("../src/graph.js");
const { addFileToJob, addFileToPart, addNoticeShare, addSeparatePart, armJob, checkJobs, createJob, createNoticeJob, createSeparateJob, getJob, setJobPasswordTo, onUnsentExpired, unsentShares } =
  await import("../src/jobs.js");
const { quietLog } = await import("./helpers.js");
const vaultwarden = await import("../src/vaultwarden.js");
const sms = await import("../src/sms.js");
const { deliveryFor } = await import("../src/deliveries.js");
const { passwordRefFor } = await import("../src/passwordRefs.js");
const { pendingNotices } = await import("../src/notices.js");
const { sentMailFor } = await import("../src/history.js");
const sendSms = vi.mocked(sms.sendSms);
const deleteSend = vi.mocked(vaultwarden.deleteSend);

const findSentMessage = vi.mocked(graph.findSentMessage);
const sendMail = vi.mocked(graph.sendMail);

const sentMessage = {
  subject: "Unterlagen",
  from: "max@example.test",
  recipients: { to: ["erika@example.test"], cc: ["chef@example.test"], bcc: [] },
};

const newJob = (fileName = "a.pdf", maxAccess?: number) =>
  createJob("oid", "max@example.test", fileName, "https://vault/#/send/x", maxAccess);
const arm = (id: string, overrides: Partial<Parameters<typeof armJob>[1]> = {}) =>
  armJob(id, { graphToken: "token", tokenExpiresAt: Date.now() + 3_600_000, armedAt: Date.now(), ...overrides });

/** Jobs stay in the module between tests – only the job of the current test is "sent". */
const sentOnlyFor = (jobId: string, message = sentMessage, mailbox?: string) =>
  findSentMessage.mockImplementation(async (_token, id, _since, box) =>
    id === jobId && box === mailbox ? message : null,
  );

beforeEach(() => {
  findSentMessage.mockReset().mockResolvedValue(null);
  sendMail.mockReset().mockResolvedValue(undefined);
  sendSms.mockReset().mockResolvedValue("77000001");
  deleteSend.mockClear();
});

describe("password jobs", () => {
  it("does nothing until the mail is armed", async () => {
    const id = await newJob();
    await checkJobs(quietLog());
    expect(findSentMessage.mock.calls.some((c) => c[1] === id)).toBe(false);
    expect(sendMail).not.toHaveBeenCalled();
  });

  it("waits while the mail is not in Sent Items yet", async () => {
    const id = await newJob();
    await arm(id);
    await checkJobs(quietLog());
    expect(findSentMessage.mock.calls.some((c) => c[1] === id)).toBe(true);
    expect(sendMail).not.toHaveBeenCalled();
    expect(getJob(id)?.armed).toBeDefined();
    expect(getJob(id)?.dispatched).toBe(false);
  });

  it("sends one password mail to the actual recipients once the mail was sent", async () => {
    const id = await newJob();
    await addFileToJob(id, "Upload-Link für Ihre Dateien");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());
    await checkJobs(quietLog()); // a second pass must not send again

    expect(sendMail).toHaveBeenCalledOnce();
    const [token, recipients, subject, html, from] = sendMail.mock.calls[0];
    expect(token).toBe("token");
    expect(recipients).toEqual(sentMessage.recipients);
    expect(subject).toBe("Kennwort für: Unterlagen");
    expect(html).toContain("die folgenden Freigaben");
    expect(html).toContain("a.pdf");
    expect(html).toContain("Upload-Link für Ihre Dateien");
    expect(html).toContain("https://vault/#/send/x");
    expect(from).toBe("max@example.test");
    expect(getJob(id)?.dispatched).toBe(true);
  });

  it("names all files of a step in its one password mail", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", maxAccessCount: 3,
    });
    await addFileToPart(id, "s1", "b.png");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    const [, , subject, html] = sendMail.mock.calls[0];
    expect(subject).toBe("Kennwort für 2 Dateien: Unterlagen");
    expect(html).toContain("a.pdf");
    expect(html).toContain("b.png");
    expect(getJob(id)?.fileNames).toEqual(["a.pdf", "b.png"]);
  });

  it("sends one password mail per step, each with its own password link", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", maxAccessCount: 3,
    });
    await addSeparatePart(id, { step: "s2", fileNames: ["b.png"], sendUrl: "https://vault/#/send/b", maxAccessCount: 1 });
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());
    await checkJobs(quietLog()); // no second round

    expect(sendMail).toHaveBeenCalledTimes(2);
    const [first, second] = sendMail.mock.calls;
    expect(first[1]).toEqual(sentMessage.recipients);
    expect(first[2]).toBe("Kennwort für „a.pdf“: Unterlagen");
    expect(first[3]).toContain("https://vault/#/send/a");
    expect(first[3]).toContain("bis zu 3 Abrufe");
    expect(first[3]).not.toContain("b.png");
    expect(second[2]).toBe("Kennwort für „b.png“: Unterlagen");
    expect(second[3]).toContain("https://vault/#/send/b");
    expect(second[3]).toContain("nur einmal abrufbar");
    expect(getJob(id)?.dispatched).toBe(true);
  });

  it("does not send a step's password mail twice when a later one failed", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", maxAccessCount: 3,
    });
    await addSeparatePart(id, { step: "s2", fileNames: ["b.png"], sendUrl: "https://vault/#/send/b", maxAccessCount: 3 });
    await arm(id);
    sentOnlyFor(id);
    sendMail.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error("Graph 503"));

    await checkJobs(quietLog()); // a.pdf sent, b.png failed
    expect(getJob(id)?.dispatched).toBe(false);
    await checkJobs(quietLog()); // retry: only b.png

    expect(sendMail.mock.calls.map((c) => c[2])).toEqual([
      "Kennwort für „a.pdf“: Unterlagen",
      "Kennwort für „b.png“: Unterlagen",
      "Kennwort für „b.png“: Unterlagen",
    ]);
    expect(getJob(id)?.dispatched).toBe(true);
  });

  it("sends the password only to the other address when one was set", async () => {
    const id = await newJob();
    await setJobPasswordTo(id, "kollege@example.test");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    const [, recipients, subject, , from] = sendMail.mock.calls[0];
    expect(recipients).toEqual({ to: ["kollege@example.test"], cc: [], bcc: [] }); // not erika / chef
    expect(subject).toBe("Kennwort für: Unterlagen");
    expect(from).toBe("max@example.test");
  });

  it("states how often the password link can be opened", async () => {
    for (const [count, text] of [[3, "bis zu 3 Abrufe"], [1, "nur einmal abrufbar"]] as const) {
      sendMail.mockClear();
      const id = await newJob("a.pdf", count);
      await arm(id);
      sentOnlyFor(id);
      await checkJobs(quietLog());
      expect(sendMail.mock.calls[0][3]).toContain(text);
    }
  });

  it("escapes file names and subjects in the password mail", async () => {
    const id = await newJob("<script>.pdf");
    await arm(id);
    sentOnlyFor(id, { ...sentMessage, subject: "A & <b>" });
    await checkJobs(quietLog());
    const html = sendMail.mock.calls[0][3];
    expect(html).not.toContain("<script>");
    expect(html).toContain("&lt;script&gt;.pdf");
    expect(html).toContain("A &amp; &lt;b&gt;");
  });

  it("also looks in the shared mailbox and sends from it", async () => {
    const id = await newJob();
    await arm(id, { mailbox: "info@example.test" });
    sentOnlyFor(id, { ...sentMessage, from: "info@example.test" }, "info@example.test");

    await checkJobs(quietLog());

    expect(findSentMessage).toHaveBeenCalledWith("token", id, expect.any(Number), "info@example.test");
    expect(sendMail).toHaveBeenCalledOnce();
    expect(sendMail.mock.calls[0][4]).toBe("info@example.test");
  });

  it("disarms but keeps the job when the mail never shows up (cancelled send)", async () => {
    const id = await newJob();
    await arm(id, { armedAt: Date.now() - 16 * 60_000 }); // older than the 15-minute window
    await checkJobs(quietLog());
    expect(findSentMessage.mock.calls.some((c) => c[1] === id)).toBe(false);
    expect(getJob(id)?.armed).toBeUndefined();
    expect(getJob(id)?.dispatched).toBe(false);

    // further files of the same mail still join this job → only one password mail later
    await addFileToJob(id, "b.pdf");
    expect(getJob(id)?.fileNames).toEqual(["a.pdf", "b.pdf"]);
  });

  it("keeps trying when Graph fails temporarily", async () => {
    const id = await newJob();
    await arm(id);
    findSentMessage.mockImplementation(async (_token, jobId) => {
      if (jobId === id) throw new Error("Graph down");
      return null;
    });
    const log = quietLog();
    await checkJobs(log);
    expect(log.warn).toHaveBeenCalled();
    expect(getJob(id)?.armed).toBeDefined();
    expect(getJob(id)?.dispatched).toBe(false);
  });
});


describe("password by SMS, deliveries and download notices", () => {
  const share = (token: string, name: string) => ({ token, name, expireDate: "2099-12-01" });
  const smsJob = async (token: string, number = "+491701234567") => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share(token, "a.pdf")],
    });
    await setJobPasswordTo(id, undefined, number);
    return id;
  };

  it("sends no SMS to a country SMS_ALLOWED_COUNTRIES no longer allows – only the note to the sender", async () => {
    const id = await smsJob("tok-sms-us", "+12025550123"); // entered before the admin limited the countries
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendSms).not.toHaveBeenCalled();
    expect(sendMail).toHaveBeenCalledOnce();
    const [, recipients, subject, html] = sendMail.mock.calls[0];
    expect(recipients).toEqual({ to: ["max@example.test"], cc: [], bcc: [] });
    expect(subject).toBe("SMS nicht versendet – Unterlagen");
    expect(html).toContain("Deutschland, Österreich und Schweiz");
    expect(html).not.toContain("KateGeteFi45");
    expect(deliveryFor("tok-sms-us", "oid")).toMatchObject({ channel: "sms", status: "failed" });
  });

  it("sends the password of a step by SMS instead of a mail and keeps the Send for sending it again", async () => {
    const id = await smsJob("tok-sms-1");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).not.toHaveBeenCalled();
    expect(sendSms).toHaveBeenCalledOnce();
    const [to, text] = sendSms.mock.calls[0];
    expect(to).toBe("+491701234567");
    expect(text).toContain("KateGeteFi45"); // read from the Send right before sending
    expect(text).toContain("Kennwort fuer \"a.pdf\""); // which link the password belongs to
    expect(deleteSend).not.toHaveBeenCalled(); // the password stays readable until the links expire
    expect(deliveryFor("tok-sms-1", "oid")).toMatchObject({ channel: "sms", to: "+49 170 •••• 67", messageId: "77000001", status: "submitted" });
    expect(passwordRefFor("tok-sms-1", "oid")).toMatchObject({ sendId: "send-a", sms: "+491701234567", subject: "Unterlagen" });
    expect(getJob(id)?.dispatched).toBe(true);
  });

  it("sends the password of one step on every way its links chose – one password mail and one SMS", async () => {
    // A file for the recipients, an upload link with the same password for them too, then a file by SMS.
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-way-1", "a.pdf")], deliveries: [{ channel: "recipients" }],
    });
    await addFileToPart(id, "s1", "Upload-Link", share("tok-way-2", "Upload-Link"), { channel: "recipients" });
    await addFileToPart(id, "s1", "b.pdf", share("tok-way-3", "b.pdf"), { channel: "sms", to: "+491701234567" });
    expect(getJob(id)?.separate?.[0].deliveries).toEqual([{ channel: "recipients" }, { channel: "sms", to: "+491701234567" }]);
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce(); // one password mail for the step …
    expect(sendMail.mock.calls[0][1]).toEqual(sentMessage.recipients);
    expect(sendSms).toHaveBeenCalledOnce(); // … and one SMS with the same password
    expect(sendSms.mock.calls[0][0]).toBe("+491701234567");
    expect(getJob(id)?.dispatched).toBe(true);
    expect(passwordRefFor("tok-way-1", "oid")).toMatchObject({ sms: "+491701234567" });

    // A second pass sends nothing again.
    await checkJobs(quietLog());
    expect(sendMail).toHaveBeenCalledOnce();
    expect(sendSms).toHaveBeenCalledOnce();
  });

  it("writes the password mail and SMS of a step in the language chosen for the recipients – notes stay in LANGUAGE", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-lang-1", "a.pdf")], language: "en",
      deliveries: [{ channel: "recipients" }, { channel: "sms", to: "+491701234567" }],
    });
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    const [, , subject, html] = sendMail.mock.calls[0];
    expect(subject).toBe("Password for “a.pdf”: Unterlagen"); // English, although LANGUAGE is de in the tests
    expect(html).toContain("Show password");
    expect(html).not.toContain("Kennwort anzeigen");
    expect(sendSms.mock.calls[0][1]).toMatch(/^Password for "a\.pdf"/);
    // Sending it again later uses the same language.
    expect(passwordRefFor("tok-lang-1", "oid")?.language).toBe("en");
  });

  it("does not send a way twice when a later one failed", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-retry-1", "a.pdf")], deliveries: [{ channel: "recipients" }, { channel: "mail", to: "kollege@example.test" }],
    });
    await arm(id);
    sentOnlyFor(id);
    sendMail.mockResolvedValueOnce(undefined).mockRejectedValueOnce(new Error("Graph down"));
    await checkJobs(quietLog()).catch(() => undefined);
    await checkJobs(quietLog()).catch(() => undefined);
    const targets = sendMail.mock.calls.map((call) => (call[1] as { to: string[] }).to[0]);
    expect(targets.filter((to) => to === "erika@example.test")).toHaveLength(1); // the recipients only once
    expect(targets.filter((to) => to === "kollege@example.test")).toHaveLength(2); // the failed one again
  });

  it("sends no password another way when the SMS fails – only a note without password to the sender", async () => {
    const id = await smsJob("tok-sms-2");
    await arm(id);
    sentOnlyFor(id);
    sendSms.mockRejectedValueOnce(new Error("seven.io: no credit"));

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    const [, recipients, subject, html] = sendMail.mock.calls[0];
    expect(recipients).toEqual({ to: ["max@example.test"], cc: [], bcc: [] });
    expect(subject).toBe("SMS nicht versendet – Unterlagen");
    expect(html).toContain("konnte nicht versendet werden.");
    expect(html).toContain("Grund: seven.io: no credit");
    expect(html).toContain("„Kennwort“");
    expect(html).not.toContain("KateGeteFi45"); // neither the password …
    expect(html).not.toContain("vault/#/send"); // … nor a link to it
    expect(deliveryFor("tok-sms-2", "oid")).toMatchObject({ channel: "sms", status: "failed" });
    expect(passwordRefFor("tok-sms-2", "oid")?.sendId).toBe("send-a"); // can be sent again
    expect(getJob(id)?.dispatched).toBe(true); // no retry of the SMS
  });

  it("records mail deliveries per link and registers the links for the download confirmation", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-mail-1", "a.pdf")],
    });
    await addFileToPart(id, "s1", "b.png", share("tok-mail-2", "b.png"));
    await addFileToPart(id, "s1", "Upload-Link für Ihre Dateien",
      { ...share("tok-mail-3", "Upload-Link für Ihre Dateien"), request: true, knownFiles: ["vorher.pdf"] });
    await arm(id);
    sentOnlyFor(id, { ...sentMessage, id: "AAMk-sent" });

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    expect(deliveryFor("tok-mail-2", "oid")).toMatchObject({ channel: "mail", to: "erika@example.test, chef@example.test" });
    expect(deliveryFor("tok-mail-2", "someone-else")).toBeUndefined();
    const notices = new Map(pendingNotices());
    expect(notices.get("tok-mail-1")).toMatchObject({ fileName: "a.pdf", subject: "Unterlagen", notifyTo: "max@example.test" });
    expect(notices.get("tok-mail-1")?.kind).toBeUndefined(); // file link: first download
    expect(notices.get("tok-mail-2")?.fileName).toBe("b.png");
    // The upload link is watched for received files (UPLOAD_NOTICE), not for downloads.
    expect(notices.get("tok-mail-3")).toMatchObject({ kind: "upload", knownFiles: ["vorher.pdf"], subject: "Unterlagen" });
    // The password can be sent again later; the mail's recipients are suggested.
    expect(passwordRefFor("tok-mail-2", "oid")).toMatchObject({
      sendId: "send-a", tokens: ["tok-mail-1", "tok-mail-2", "tok-mail-3"], fileNames: ["a.pdf", "b.png", "Upload-Link für Ihre Dateien"],
      recipients: { to: ["erika@example.test"], cc: ["chef@example.test"], bcc: [] }, maxAccessCount: 3,
    });
    expect(passwordRefFor("tok-mail-2", "someone-else")).toBeNull();
    // "Verlauf": every link knows its mail ' subject, and the ID to open it.
    expect(sentMailFor("tok-mail-3", "oid")).toMatchObject({ subject: "Unterlagen", mailId: "AAMk-sent" });
    expect(sentMailFor("tok-mail-3", "someone-else")).toBeUndefined();
  });

  it("keeps no mail ID for a mail found in a shared mailbox – the pane can only open the user's own mails", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-shared-1", "a.pdf")],
    });
    await arm(id, { mailbox: "info@example.test" });
    sentOnlyFor(id, { ...sentMessage, id: "AAMk-info", from: "info@example.test" }, "info@example.test");

    await checkJobs(quietLog());

    expect(sentMailFor("tok-shared-1", "oid")).toMatchObject({ subject: "Unterlagen" });
    expect(sentMailFor("tok-shared-1", "oid")?.mailId).toBeUndefined();
  });

  it("sends the password mail with the distribution of the sent mail – To, Cc and Bcc as they were", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-bcc", maxAccessCount: 3,
      shares: [share("tok-bcc", "a.pdf")],
    });
    await arm(id);
    const distribution = { to: ["erika@example.test"], cc: ["chef@example.test"], bcc: ["archiv@example.test", "buergerin@example.test"] };
    sentOnlyFor(id, { ...sentMessage, recipients: distribution });

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    expect(sendMail.mock.calls[0][1]).toEqual(distribution); // Bcc recipients get it too, still hidden from the others
    expect(deliveryFor("tok-bcc", "oid")?.to).toBe("erika@example.test, chef@example.test"); // Bcc not shown
    expect(passwordRefFor("tok-bcc", "oid")?.recipients).toEqual(distribution); // sent again exactly like this
  });

  it("remembers another address as such, for sending the password again the same way", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-other", maxAccessCount: 3,
      shares: [share("tok-other", "a.pdf")],
    });
    await setJobPasswordTo(id, "kollege@example.test");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail.mock.calls[0][1]).toEqual({ to: ["kollege@example.test"], cc: [], bcc: [] });
    expect(passwordRefFor("tok-other", "oid")).toMatchObject({ otherTo: ["kollege@example.test"], recipients: sentMessage.recipients });
  });

  it("links without password: no password mail, no SMS – only the download confirmation once the mail was sent", async () => {
    const id = await createNoticeJob("oid", "max@example.test", share("tok-free-1", "frei.pdf"));
    await addNoticeShare(id, share("tok-free-2", "frei2.pdf"));
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).not.toHaveBeenCalled();
    expect(getJob(id)?.dispatched).toBe(true);
    const notices = new Map(pendingNotices());
    expect(notices.get("tok-free-1")).toMatchObject({ fileName: "frei.pdf", subject: "Unterlagen" });
    expect(notices.get("tok-free-2")?.fileName).toBe("frei2.pdf");
  });
});

describe("mails that were never sent", () => {
  const share = (token: string) => ({ token, name: `${token}.pdf`, expireDate: "2099-12-01" });
  const DAY = 86_400_000;
  /** Runs the watcher as if `days` had passed – every job created so far is then expired. */
  const checkAfter = async (days: number, log = quietLog()) => {
    const now = Date.now();
    const clock = vi.spyOn(Date, "now").mockReturnValue(now + days * DAY);
    try {
      await checkJobs(log);
    } finally {
      clock.mockRestore();
    }
    return log;
  };
  const handler = vi.fn<(links: UnsentLinks, log: unknown) => Promise<void>>();

  beforeEach(() => {
    handler.mockReset().mockResolvedValue(undefined);
    onUnsentExpired(handler);
  });
  afterEach(() => onUnsentExpired(undefined));

  it("lists the links of a draft not sent yet, with the day they are revoked", async () => {
    const id = await createSeparateJob("oid-unsent", "max@example.test", {
      step: "s1", fileNames: ["u1.pdf"], sendUrl: "https://vault/#/send/u1", sendId: "send-u1", maxAccessCount: 3, shares: [share("tok-u1")],
    });
    const unsent = unsentShares("oid-unsent");
    expect(unsent.get("tok-u1")).toBe(getJob(id)!.expiresAt);
    expect(unsentShares("someone-else").size).toBe(0);

    await arm(id); // "Send" pressed: on its way
    expect(unsentShares("oid-unsent").has("tok-u1")).toBe(false);
  });

  it("keeps a draft alive while links are added to it", async () => {
    const id = await createSeparateJob("oid-unsent", "max@example.test", {
      step: "s1", fileNames: ["u2.pdf"], sendUrl: "https://vault/#/send/u2", maxAccessCount: 3, shares: [share("tok-u2")],
    });
    getJob(id)!.expiresAt = Date.now() + 60_000; // created long ago
    await addFileToPart(id, "s1", "u3.pdf", share("tok-u3"));
    expect(getJob(id)!.expiresAt).toBeGreaterThan(Date.now() + 6 * DAY);
  });

  it("revokes the links and Sends of a draft that expired without Send ever being pressed", async () => {
    const id = await createSeparateJob("oid-unsent", "max@example.test", {
      step: "s1", fileNames: ["u4.pdf"], sendUrl: "https://vault/#/send/u4", sendId: "send-u4", maxAccessCount: 3, shares: [share("tok-u4")],
    });
    await addSeparatePart(id, {
      step: "s2", fileNames: ["u5.pdf"], sendUrl: "https://vault/#/send/u5", sendId: "send-u5", maxAccessCount: 1, shares: [share("tok-u5")],
    });

    await checkAfter(6); // not expired yet
    expect(handler.mock.calls.some(([links]) => links.tokens.includes("tok-u4"))).toBe(false);

    await checkAfter(8);
    const call = handler.mock.calls.find(([links]) => links.tokens.includes("tok-u4"));
    expect(call?.[0]).toEqual({ owner: "oid-unsent", ownerUpn: "max@example.test", tokens: ["tok-u4", "tok-u5"], sendIds: ["send-u4", "send-u5"] });
    expect(getJob(id)).toBeUndefined();
  });

  it("leaves the links alone if Send was pressed (the mail may have gone out unseen) or they have no password", async () => {
    const armed = await createSeparateJob("oid-unsent", "max@example.test", {
      step: "s1", fileNames: ["u6.pdf"], sendUrl: "https://vault/#/send/u6", sendId: "send-u6", maxAccessCount: 3, shares: [share("tok-u6")],
    });
    await arm(armed, { armedAt: Date.now() - 16 * 60_000 });
    await checkJobs(quietLog()); // not found in Sent Items → disarmed, but it was armed once
    const free = await createNoticeJob("oid-unsent", "max@example.test", share("tok-u7"));

    await checkAfter(8);
    const tokens = handler.mock.calls.flatMap(([links]) => links.tokens);
    expect(tokens).not.toContain("tok-u6");
    expect(tokens).not.toContain("tok-u7");
    expect(getJob(armed)).toBeUndefined();
    expect(getJob(free)).toBeUndefined();
  });

  it("tries again later when the links could not be revoked", async () => {
    const id = await createSeparateJob("oid-unsent", "max@example.test", {
      step: "s1", fileNames: ["u8.pdf"], sendUrl: "https://vault/#/send/u8", sendId: "send-u8", maxAccessCount: 3, shares: [share("tok-u8")],
    });
    handler.mockRejectedValue(new Error("Nextcloud unreachable"));
    const log = await checkAfter(8);
    expect(getJob(id)).toBeDefined();
    expect(log.warn).toHaveBeenCalled();

    handler.mockReset().mockResolvedValue(undefined);
    await checkAfter(9);
    expect(handler.mock.calls.some(([links]) => links.tokens.includes("tok-u8"))).toBe(true);
    expect(getJob(id)).toBeUndefined();
  });
});
