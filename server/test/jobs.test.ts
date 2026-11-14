import { beforeEach, describe, expect, it, vi } from "vitest";

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

const graph = await import("../src/graph.js");
const { addFileToJob, addFileToPart, addNoticeShare, addSeparatePart, armJob, checkJobs, createJob, createNoticeJob, createSeparateJob, getJob, setJobPasswordTo } =
  await import("../src/jobs.js");
const { quietLog } = await import("./helpers.js");
const vaultwarden = await import("../src/vaultwarden.js");
const sms = await import("../src/sms.js");
const { deliveryFor } = await import("../src/deliveries.js");
const { pendingNotices } = await import("../src/notices.js");
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
  const smsJob = async (token: string) => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share(token, "a.pdf")],
    });
    await setJobPasswordTo(id, undefined, "+491701234567");
    return id;
  };

  it("sends the password of a step by SMS instead of a mail and deletes the Send", async () => {
    const id = await smsJob("tok-sms-1");
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).not.toHaveBeenCalled();
    expect(sendSms).toHaveBeenCalledOnce();
    const [to, text] = sendSms.mock.calls[0];
    expect(to).toBe("+491701234567");
    expect(text).toContain("KateGeteFi45"); // read from the Send right before sending
    expect(deleteSend).toHaveBeenCalledWith("send-a");
    expect(deliveryFor("tok-sms-1", "oid")).toMatchObject({ channel: "sms", to: "+49 170 •••• 67", messageId: "77000001", status: "submitted" });
    expect(getJob(id)?.dispatched).toBe(true);
  });

  it("sends the password mail to the sender when the SMS fails – never silently to the recipients", async () => {
    const id = await smsJob("tok-sms-2");
    await arm(id);
    sentOnlyFor(id);
    sendSms.mockRejectedValueOnce(new Error("seven.io: no credit"));

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    const [, recipients, subject] = sendMail.mock.calls[0];
    expect(recipients).toEqual({ to: ["max@example.test"], cc: [], bcc: [] });
    expect(subject).toMatch(/^SMS nicht versendet – Kennwort für „a\.pdf“/);
    expect(deleteSend).not.toHaveBeenCalled(); // the sender's mail still needs the Send
    expect(deliveryFor("tok-sms-2", "oid")).toMatchObject({ channel: "sms", status: "failed" });
  });

  it("records mail deliveries per link and registers the links for the download confirmation", async () => {
    const id = await createSeparateJob("oid", "max@example.test", {
      step: "s1", fileNames: ["a.pdf"], sendUrl: "https://vault/#/send/a", sendId: "send-a", maxAccessCount: 3,
      shares: [share("tok-mail-1", "a.pdf")],
    });
    await addFileToPart(id, "s1", "b.png", share("tok-mail-2", "b.png"));
    await arm(id);
    sentOnlyFor(id);

    await checkJobs(quietLog());

    expect(sendMail).toHaveBeenCalledOnce();
    expect(deliveryFor("tok-mail-2", "oid")).toMatchObject({ channel: "mail", to: "erika@example.test, chef@example.test" });
    expect(deliveryFor("tok-mail-2", "someone-else")).toBeUndefined();
    const notices = new Map(pendingNotices());
    expect(notices.get("tok-mail-1")).toMatchObject({ fileName: "a.pdf", subject: "Unterlagen", notifyTo: "max@example.test" });
    expect(notices.get("tok-mail-2")?.fileName).toBe("b.png");
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
