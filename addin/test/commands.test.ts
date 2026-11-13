/**
 * The Outlook event handlers (src/commands/commands.ts): attachment policy on send, arming the password mail,
 * the info bar while composing – the most critical behaviour, since a mistake here can block sending mail.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";
import { installOfficeMock, type Attachment } from "./officeMock";

const office = installOfficeMock();
await import("../src/commands/commands");
const { resetTokenCache, TIMEOUTS } = await import("../src/shared/office");

const MB = 1024 * 1024;
const big: Attachment = { id: "a1", name: "Fotos.zip", size: 5 * MB, attachmentType: "file", isInline: false };
const small: Attachment = { id: "a2", name: "Notiz.txt", size: 10_000, attachmentType: "file", isInline: false };

interface Server {
  policy: "enforce" | "suggest";
  reachable: boolean;
  armStatus: number;
  trace?: boolean;
}
let server: Server;
const requests: Array<{ url: string; body: unknown }> = [];

function installFetch() {
  vi.stubGlobal(
    "fetch",
    vi.fn(async (url: string, init?: RequestInit) => {
      requests.push({ url, body: init?.body ? JSON.parse(String(init.body)) : undefined });
      if (!server.reachable) throw new TypeError("Failed to fetch");
      if (url.endsWith("/api/config")) {
        return new Response(
          JSON.stringify({ thresholdMb: 1, attachmentPolicy: server.policy, defaultExpireDays: 14, maxExpireDays: 90, chunkSize: 10 * MB, mailFooter: "", trace: server.trace === true }),
        );
      }
      if (url.includes("/arm")) {
        return server.armStatus === 200
          ? new Response(JSON.stringify({ ok: true }))
          : new Response(JSON.stringify({ error: "Graph nicht erreichbar" }), { status: server.armStatus });
      }
      throw new Error(`unexpected request ${url}`);
    }),
  );
}

const draft = (state: Record<string, unknown>) => {
  office.state.session["cloudshare-state"] = JSON.stringify(state);
};

/** Runs the send handler and returns the options passed to event.completed. */
async function send() {
  const completed = vi.fn();
  await office.handlers.onMessageSendHandler({ completed });
  expect(completed).toHaveBeenCalledOnce();
  const options = completed.mock.calls[0][0] as Record<string, unknown>;
  // Outlook on the web rejects longer values and leaves the send dialog spinning.
  expect(String(options.cancelLabel ?? "").length).toBeLessThanOrEqual(20);
  expect(String(options.errorMessage ?? "").length).toBeLessThanOrEqual(500);
  return options;
}

const armed = () => requests.filter((r) => r.url.includes("/arm"));

beforeEach(() => {
  office.reset();
  resetTokenCache();
  server = { policy: "enforce", reachable: true, armStatus: 200 };
  requests.length = 0;
  installFetch();
});

describe("send – policy enforce", () => {
  it("blocks large attachments and offers the CloudShare button (Mailbox 1.14)", async () => {
    office.state.attachments = [big];
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toMatch(/dürfen nur mit CloudShare/);
    expect(result.errorMessage).toContain("Fotos.zip");
    expect(result.cancelLabel).toBe("CloudShare öffnen");
    expect(result.commandId).toBe("cloudshareOpenPane");
    expect(result.sendModeOverride).toBeUndefined(); // no "Send anyway"
  });

  it("stays within Outlook's 500 characters for many long attachment names", async () => {
    office.state.attachments = Array.from({ length: 20 }, (_, i) => ({ ...big, id: `a${i}`, name: `${"x".repeat(80)}-${i}.zip` }));
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toMatch(/dürfen nur mit CloudShare/);
  });

  it("points older clients to the ribbon button", async () => {
    office.state.attachments = [big];
    office.state.supported = { Mailbox: "1.12" };
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toContain("„CloudShare“ in der Menüleiste");
    expect(result.cancelLabel).toBeUndefined();
  });

  it("does not arm the password mail while the mail is still blocked", async () => {
    office.state.attachments = [big];
    draft({ jobId: "job-1" });
    await send();
    expect(armed()).toHaveLength(0);
  });

  it("lets small, inline and non-file attachments through", async () => {
    office.state.attachments = [
      small,
      { ...big, id: "i", isInline: true },
      { ...big, id: "m", attachmentType: "item" },
      { ...big, id: "edge", size: 1 * MB }, // exactly at the threshold is not "larger"
    ];
    expect(await send()).toEqual({ allowEvent: true });
  });
});

describe("send – policy suggest", () => {
  beforeEach(() => {
    server.policy = "suggest";
  });

  it("arms the password mail first, then offers 'Send anyway'", async () => {
    office.state.attachments = [big];
    draft({ jobId: "job-1" });
    const result = await send();
    expect(armed()).toHaveLength(1);
    expect(result.allowEvent).toBe(false);
    expect(result.sendModeOverride).toBe("promptUser");
    expect(result.cancelLabel).toBe("CloudShare öffnen");
  });

  it("does not block clients that cannot offer 'Send anyway'", async () => {
    office.state.attachments = [big];
    office.state.supported = { Mailbox: "1.12" };
    expect(await send()).toEqual({ allowEvent: true });
  });
});

describe("send – password mail", () => {
  it("arms the job, tags the draft and allows sending", async () => {
    draft({ jobId: "job-1" });
    expect(await send()).toEqual({ allowEvent: true });
    expect(armed()).toHaveLength(1);
    expect(armed()[0].url).toContain("/api/jobs/job-1/arm");
    expect(armed()[0].body).toEqual({});
    expect(office.state.customProps.cloudshareJob).toBe("job-1");
  });

  it("still arms the job while one of the inserted links is in the mail", async () => {
    draft({ jobId: "job-1", links: ["tokA", "tokB"] });
    office.state.body = '<a href="https://cloud.example.test/s/tokB">Herunterladen</a>';
    expect(await send()).toEqual({ allowEvent: true });
    expect(armed()).toHaveLength(1);
  });

  it("sends no password mail when all inserted links were deleted from the mail", async () => {
    draft({ jobId: "job-1", password: "geheim", links: ["tokA"] });
    office.state.customProps.cloudshareJob = "job-1";
    office.state.body = "<p>Nur Text, kein Link mehr</p>";
    expect(await send()).toEqual({ allowEvent: true });
    expect(armed()).toHaveLength(0);
    expect(office.state.customProps.cloudshareJob).toBeUndefined();
    expect(JSON.parse(office.state.session["cloudshare-state"])).toEqual({});
  });

  it("never releases drafts without recorded links (inserted before this version)", async () => {
    draft({ jobId: "job-1" });
    office.state.body = "";
    await send();
    expect(armed()).toHaveLength(1);
  });

  it("reports a shared mailbox in the From field", async () => {
    draft({ jobId: "job-1" });
    office.state.from = "Info@Example.test";
    await send();
    expect(armed()[0].body).toEqual({ from: "Info@Example.test" });
  });

  it("treats the own address in another spelling as the own mailbox", async () => {
    draft({ jobId: "job-1" });
    office.state.from = "MAX@example.test";
    await send();
    expect(armed()[0].body).toEqual({});
  });

  it("blocks sending when the password mail cannot be prepared", async () => {
    draft({ jobId: "job-1" });
    server.armStatus = 502;
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toMatch(/nicht vorbereiten.*Graph nicht erreichbar/);
  });

  it("does not hang when the Microsoft sign-in never answers – names the step instead", async () => {
    draft({ jobId: "job-1" });
    const auth = (globalThis as unknown as { Office: { auth: { getAccessToken: ReturnType<typeof vi.fn> } } }).Office.auth;
    auth.getAccessToken.mockImplementationOnce(() => new Promise(() => undefined)); // never settles
    const saved = TIMEOUTS.token;
    TIMEOUTS.token = 50;
    try {
      const result = await send();
      expect(result.allowEvent).toBe(false);
      expect(result.errorMessage).toContain("Zeitüberschreitung: Die Anmeldung bei Microsoft (SSO)");
      expect(armed()).toHaveLength(0);
    } finally {
      TIMEOUTS.token = saved;
    }
  });

  it("shows Office SSO errors with a readable text and their code (not [object Object])", async () => {
    draft({ jobId: "job-1" });
    const auth = (globalThis as unknown as { Office: { auth: { getAccessToken: ReturnType<typeof vi.fn> } } }).Office.auth;
    // Office rejects with a plain object, not an Error.
    auth.getAccessToken.mockRejectedValueOnce({ code: 13007, name: "Error", message: "An unexpected error occurred" });
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toContain("kein Anmeldetoken für CloudShare");
    expect(result.errorMessage).toContain("(Code 13007)");
    expect(result.errorMessage).not.toContain("[object Object]");
  });
});

describe("send – diagnostics (CLIENT_TRACE)", () => {
  const traces = () =>
    requests.filter((r) => r.url.includes("/api/trace?")).map((r) => new URLSearchParams(r.url.split("?")[1]).get("step"));

  it("stays silent unless the server switched tracing on", async () => {
    draft({ jobId: "job-1" });
    await send();
    expect(traces()).toEqual([]);
  });

  it("reports every step of the send handler once tracing is on", async () => {
    office.state.roaming["cloudshare-config"] = { trace: true }; // cached from an earlier /api/config
    server.trace = true;
    draft({ jobId: "job-1" });
    expect(await send()).toEqual({ allowEvent: true });
    expect(traces()).toEqual([
      "start", "state", "config", "attachments", "arm:tag", "arm:from", "arm:token", "arm:server", "arm:done", "completed",
    ]);
  });
});

describe("send – server unreachable", () => {
  it("lets mail through when no configuration was ever loaded", async () => {
    server.reachable = false;
    office.state.attachments = [big];
    expect(await send()).toEqual({ allowEvent: true });
  });

  it("keeps enforcing with the cached configuration", async () => {
    office.state.attachments = [big];
    await send(); // loads and caches the configuration
    server.reachable = false;
    const result = await send();
    expect(result.allowEvent).toBe(false);
    expect(result.errorMessage).toMatch(/dürfen nur mit CloudShare/);
  });
});

describe("info bar while composing", () => {
  const changed = async () => {
    const completed = vi.fn();
    await office.handlers.onAttachmentsChangedHandler({ completed });
    expect(completed).toHaveBeenCalledOnce();
  };

  it("offers CloudShare for large attachments", async () => {
    office.state.attachments = [big];
    await changed();
    const [bar] = office.state.notifications.replaced;
    expect(bar.type).toBe("insightMessage");
    expect(String(bar.message)).toContain("„Fotos.zip“ ist größer als 1 MB");
    expect(String(bar.message).length).toBeLessThanOrEqual(150);
    expect(bar.actions).toEqual([expect.objectContaining({ actionType: "showTaskPane", commandId: "cloudshareOpenPane" })]);
  });

  it("keeps the message within Outlook's 150 characters for long names", async () => {
    office.state.attachments = [{ ...big, name: `${"x".repeat(200)}.zip` }];
    await changed();
    expect(String(office.state.notifications.replaced[0].message).length).toBeLessThanOrEqual(150);
  });

  it("removes the bar when no large attachment is left", async () => {
    office.state.attachments = [small];
    await changed();
    expect(office.state.notifications.replaced).toHaveLength(0);
    expect(office.state.notifications.removed).toEqual(["cloudshare-large-attachment"]);
  });

  it("also offers CloudShare under 'suggest' (asking a question, not blocking)", async () => {
    server.policy = "suggest";
    office.state.attachments = [big];
    await changed();
    expect(String(office.state.notifications.replaced[0].message)).toContain("Als kennwortgeschützten Nextcloud-Link senden?");
  });

  it("never fails while composing, even if the server is down", async () => {
    server.reachable = false;
    office.state.attachments = [big];
    await changed();
    expect(office.state.notifications.replaced).toHaveLength(0);
  });
});
