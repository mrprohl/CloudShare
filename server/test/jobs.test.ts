import { beforeEach, describe, expect, it, vi } from "vitest";

// The watcher talks to Graph only through these two functions – replace them with test doubles.
vi.mock("../src/graph.js", () => ({
  findSentMessage: vi.fn(),
  sendMail: vi.fn(),
}));

const graph = await import("../src/graph.js");
const { addFileToJob, armJob, checkJobs, createJob, getJob } = await import("../src/jobs.js");
const { quietLog } = await import("./helpers.js");

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

