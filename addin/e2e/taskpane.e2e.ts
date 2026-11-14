/**
 * UI tests of the taskpane (built add-in in Chromium at taskpane width). Office.js is replaced by e2e/office-mock.js,
 * the backend by route mocks below – the test checks what the user sees and what the pane does to the mail.
 */
import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page, type Route } from "@playwright/test";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const officeMock = readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), "office-mock.js"), "utf8");
const MiB = 1024 * 1024;
const STRONG = "kT7#pQ2m!xV9wR4e_Hn3";

interface Attachment { id: string; name: string; size: number; content: string }

/** Attachment with real content (Base64) of the given size. */
const attachment = (id: string, name: string, size: number): Attachment => ({
  id,
  name,
  size,
  content: Buffer.alloc(size, 0x41).toString("base64"),
});

interface Scenario {
  attachments?: Attachment[];
  theme?: string;
  bodyType?: "html" | "text";
  connected?: boolean;
  config?: Record<string, unknown> | "down";
  shares?: unknown[];
  /** New Outlook: cursor never in the body → setSelectedDataAsync inserts nothing; bodyRejectsInsert: nothing works. */
  cursorOutsideBody?: boolean;
  bodyRejectsInsert?: boolean;
  /** Office SSO failure (plain object as Office delivers it). */
  tokenError?: { code: number; name: string; message: string };
  /** Error response for POST /api/uploads, e.g. Nextcloud connection revoked. */
  uploadError?: { status: number; error: string; code?: string };
  /** Nextcloud enforces passwords for public links. */
  passwordEnforced?: boolean;
}

interface ApiCall { method: string; path: string; body: Buffer | null }

const CONFIG = {
  thresholdMb: 1,
  attachmentPolicy: "enforce",
  defaultExpireDays: 14,
  maxExpireDays: 90,
  chunkSize: 1 * MiB,
  sendMaxAccessCount: 3,
  sendMaxAccessLimit: 20,
  mailFooter: "{displayName} · {jobTitle}",
};

/** Opens the taskpane with the given mail and backend state; returns the list of API calls made by the pane. */
async function openTaskpane(page: Page, scenario: Scenario = {}): Promise<ApiCall[]> {
  const calls: ApiCall[] = [];
  let connected = scenario.connected ?? true;
  const uploads = new Map<string, { fileName: string; size: number }>();

  await page.addInitScript((cfg) => {
    (window as unknown as { __officeConfig: unknown }).__officeConfig = cfg;
  }, { attachments: scenario.attachments ?? [], theme: scenario.theme, bodyType: scenario.bodyType, tokenError: scenario.tokenError,
    cursorOutsideBody: scenario.cursorOutsideBody, bodyRejectsInsert: scenario.bodyRejectsInsert });
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "text/javascript", body: officeMock }),
  );

  await page.route("**/api/**", async (route: Route) => {
    const request = route.request();
    const url = new URL(request.url());
    const call = { method: request.method(), path: url.pathname, body: request.postDataBuffer() };
    calls.push(call);
    const json = (body: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) });
    const route_ = `${call.method} ${call.path}`;

    if (route_ === "GET /api/config") {
      return scenario.config === "down" ? json({ error: "down" }, 503) : json({ ...CONFIG, ...scenario.config });
    }
    if (route_ === "GET /api/profile") {
      return json({ displayName: "Max Mustermann", givenName: "Max", surname: "Mustermann", jobTitle: "Projektleitung",
        department: "", company: "Beispiel GmbH", email: "max.mustermann@example.com", phone: "", mobile: "", office: "" });
    }
    if (route_ === "GET /api/nextcloud/status") {
      return json(connected
        ? { connected: true, loginName: "max", freeBytes: 12.5 * 1024 ** 3, usedBytes: 37.5 * 1024 ** 3, totalBytes: 50 * 1024 ** 3,
            passwordEnforced: scenario.passwordEnforced ?? false }
        : { connected: false });
    }
    if (route_ === "POST /api/nextcloud/connect") {
      return json({ connectId: "c1", loginUrl: "https://cloud.example.com/login/v2/flow/abc" });
    }
    if (route_ === "POST /api/nextcloud/connect/c1/poll") {
      connected = true;
      return json({ connected: true, loginName: "max" });
    }
    if (route_ === "POST /api/uploads") {
      if (scenario.uploadError) return json(scenario.uploadError, scenario.uploadError.status);
      const { fileName, size } = JSON.parse(call.body!.toString());
      const uploadId = `u${uploads.size + 1}`;
      uploads.set(uploadId, { fileName, size });
      return json({ uploadId, chunkSize: 1 * MiB, totalChunks: Math.ceil(size / MiB) });
    }
    if (/^PUT \/api\/uploads\/u\d+\/chunks\/\d+$/.test(route_)) return json({ received: 1 });
    if (/^POST \/api\/uploads\/u\d+\/complete$/.test(route_)) {
      const id = call.path.split("/")[3];
      const { fileName, size } = uploads.get(id)!;
      const jobId = JSON.parse(call.body!.toString()).noPassword === true ? null : "job-1";
      return json({ shareUrl: `https://cloud.example.com/s/${id}`, expireDate: "2026-12-01", jobId, fileName, size });
    }
    if (route_ === "POST /api/requests") {
      return json({ shareUrl: "https://cloud.example.com/s/request", expireDate: "2026-12-01", jobId: "job-1" });
    }
    if (route_ === "GET /api/shares") return json(scenario.shares ?? []);
    if (/^DELETE \/api\/shares\/.+$/.test(route_)) return json({ ok: true });
    if (/^PATCH \/api\/jobs\/[^/]+$/.test(route_)) return json({ maxAccess: JSON.parse(call.body!.toString()).maxAccess });
    if (/^POST \/api\/shares\/\d+\/link$/.test(route_)) {
      const id = call.path.split("/")[3];
      const share = (scenario.shares as Array<{ id: string; kind: string; name: string }> | undefined)?.find((s) => s.id === id);
      const jobId = JSON.parse(call.body!.toString()).noPassword === true ? null : "job-1";
      return json({ kind: share?.kind ?? "file", shareUrl: `https://cloud.example.com/s/new-${id}`, expireDate: "2026-12-01",
        jobId, fileName: share?.name ?? "x", size: share?.kind === "request" ? null : 3 * MiB });
    }
    return json({ error: `not mocked: ${route_}` }, 404);
  });

  await page.goto("/taskpane.html");
  return calls;
}

const officeLog = (page: Page) =>
  page.evaluate(() => (window as unknown as { __officeLog: {
    inserted: { data: string; coercionType: string; at: "cursor" | "top" }[];
    removed: string[];
    customProps: Record<string, string>;
    session: Record<string, string>;
    openedUrls: string[];
  } }).__officeLog);

/** No serious or critical WCAG 2 A/AA violations in the current view. */
async function expectAccessible(page: Page) {
  const result = await new AxeBuilder({ page }).withTags(["wcag2a", "wcag2aa"]).analyze();
  const serious = result.violations.filter((v) => v.impact === "serious" || v.impact === "critical");
  expect(serious.map((v) => `${v.id}: ${v.nodes.map((n) => n.target.join(" ")).join(", ")}`)).toEqual([]);
}

/** The pane must not scroll sideways at taskpane width. */
async function expectNoHorizontalScroll(page: Page) {
  const overflow = await page.evaluate(() => document.documentElement.scrollWidth - document.documentElement.clientWidth);
  expect(overflow).toBeLessThanOrEqual(0);
}

const shot = async (page: Page, name: string) =>
  test.info().attach(name, { body: await page.screenshot({ fullPage: true }), contentType: "image/png" });

// --- Connection --------------------------------------------------------------------------------------------------

test("not connected: shows the connect view and connects via the Nextcloud login", async ({ page }) => {
  const calls = await openTaskpane(page, { connected: false });
  await expect(page.locator("#connect-view")).toBeVisible();
  await expect(page.locator("#send-view")).toBeHidden();
  await expectAccessible(page);
  await shot(page, "connect-view");

  await page.click("#connect");
  await expect(page.locator("#connect-hint")).toBeVisible();
  expect((await officeLog(page)).openedUrls).toEqual(["https://cloud.example.com/login/v2/flow/abc"]);

  await expect(page.locator("#send-view")).toBeVisible({ timeout: 10_000 });
  await expect(page.locator("#status")).toContainText("Die Verbindung zu Nextcloud wurde hergestellt.");
  await expect(page.locator("#account")).toContainText("max");
  expect(calls.some((c) => c.path === "/api/nextcloud/connect/c1/poll")).toBe(true);
});

test("opening the pane: profile and connection status share one sign-in (no 13008)", async ({ page }) => {
  // The footer template has placeholders → profile and Nextcloud status both need a token right at startup.
  const calls = await openTaskpane(page);
  await expect(page.locator("#send-view")).toBeVisible();
  await expect.poll(() => calls.some((c) => c.path === "/api/profile")).toBe(true);
  await expect(page.locator("#status")).toBeHidden();
  const log = await page.evaluate(() => (window as unknown as { __officeLog: { tokenRequests: number } }).__officeLog);
  expect(log.tokenRequests).toBe(1);
});

test("Microsoft sign-in fails: readable message with the Office code", async ({ page }) => {
  await openTaskpane(page, { tokenError: { code: 13007, name: "Error", message: "An unexpected error occurred" } });
  const status = page.locator("#status");
  await expect(status).toContainText("Die Anmeldung ist fehlgeschlagen:");
  await expect(status).toContainText("kein Anmeldetoken für CloudShare");
  await expect(status).toContainText("(Code 13007)");
  await expect(status).not.toContainText("[object Object]");
  await shot(page, "sso-error");
});

test("server unreachable: says so instead of showing an empty pane", async ({ page }) => {
  await openTaskpane(page, { config: "down" });
  await expect(page.locator("#status")).toContainText("CloudShare ist derzeit nicht erreichbar");
});

// --- Sending files -----------------------------------------------------------------------------------------------

test("preselects large attachments and proposes a random password", async ({ page }) => {
  await openTaskpane(page, {
    attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB), attachment("a2", "Notiz.txt", 200 * 1024)],
  });
  const rows = page.locator("#attachments label.list-item");
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0).locator("input")).toBeChecked();
  await expect(rows.nth(0).locator(".badge")).toHaveText("Groß");
  await expect(rows.nth(1).locator("input")).not.toBeChecked();
  await expect(rows.nth(1).locator(".badge")).toHaveCount(0);

  const password = page.locator("#password");
  await expect(password).toHaveAttribute("readonly", "");
  expect(await password.inputValue()).toMatch(/^([A-Z][aeiou][a-z][aeiou]){2}[A-Z][aeiou][2-9]{2}$/); // pronounceable
  await expect(page.locator("#max-access")).toHaveValue("3");
  await expect(page.locator("#expire")).toHaveValue("14");
  await expect(page.locator("#account")).toHaveText("max [12,5 GB von 50 GB frei]");

  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "send-light");
});

test("uploads a large attachment in chunks, removes it and inserts the card", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)] });
  const password = await page.locator("#password").inputValue();
  await page.click("#submit");

  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const chunks = calls.filter((c) => c.method === "PUT");
  expect(chunks.map((c) => c.body!.length)).toEqual([MiB, MiB, MiB / 2]);
  expect(Buffer.concat(chunks.map((c) => c.body!)).equals(Buffer.alloc(2.5 * MiB, 0x41))).toBe(true);

  const complete = calls.find((c) => c.path.endsWith("/complete"))!;
  expect(JSON.parse(complete.body!.toString())).toMatchObject({ password, expireDays: 14, maxAccess: 3 });

  const log = await officeLog(page);
  expect(log.removed).toEqual(["a1"]);
  expect(log.inserted).toHaveLength(1);
  expect(log.inserted[0].coercionType).toBe("html");
  // Empty line after the card: the cursor ends up below it, so a second card does not land inside this one.
  expect(log.inserted[0].data.endsWith("<p style=\"margin:0;\"><br></p>")).toBe(true);
  expect(log.inserted[0].data).toContain("https://cloud.example.com/s/u1");
  expect(log.inserted[0].data).toContain("Präsentation.pptx");
  expect(log.inserted[0].data).toContain("Max Mustermann · Projektleitung"); // footer from the profile
  expect(log.customProps.cloudshareJob).toBe("job-1"); // links the draft to the password mail
  expect(JSON.parse(log.session["cloudshare-state"])).toMatchObject({ jobId: "job-1", separate: true });

  // The next "Als Link einfügen" gets its own password and password mail; only "Ohne Kennwort" is fixed now.
  await expect(page.locator("#mode-random")).toBeEnabled();
  await expect(page.locator("#mode-none")).toBeDisabled();
  await expect(page.locator("#password")).not.toHaveValue(password);
  await expect(page.locator("#password-hint")).toContainText("Der nächste Link erhält ein neues Kennwort");
  await shot(page, "after-upload");

  // "Kennwort abrufbar" applies to the next link.
  await page.fill("#max-access", "7");
  await page.press("#max-access", "Tab");
  await expect(page.locator("#status")).toContainText("7-mal abrufbar (für den nächsten Link)");
  expect(calls.some((c) => c.method === "PATCH")).toBe(false);
});

test("Link gültig für / Kennwort abrufbar: capped at the maximum with a note, confirmed otherwise", async ({ page }) => {
  await openTaskpane(page);
  const status = page.locator("#status");

  await page.fill("#expire", "500");
  await page.press("#expire", "Tab");
  await expect(page.locator("#expire")).toHaveValue("90");
  await expect(status).toContainText("Höchstens 90 Tage möglich – der Wert wurde auf 90 gesetzt.");
  await expect(status).toContainText("Neue Links sind 90 Tage gültig.");

  await page.fill("#expire", "10");
  await page.press("#expire", "Tab");
  await expect(status).toHaveText("Neue Links sind 10 Tage gültig.");

  await page.fill("#max-access", "50");
  await page.press("#max-access", "Tab");
  await expect(page.locator("#max-access")).toHaveValue("20");
  await expect(status).toContainText("Höchstens 20 Abrufe möglich – der Wert wurde auf 20 gesetzt.");

  await page.fill("#max-access", "1");
  await page.press("#max-access", "Tab");
  await expect(status).toHaveText("Der Link zum Kennwort ist einmal abrufbar.");
});

test("every 'Als Link einfügen' gets its own password; files of one click share it", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Angebot.pdf", 1.5 * MiB), attachment("a2", "Plan.zip", 1.5 * MiB)],
  });
  await page.click("#submit"); // click 1: both files together
  await expect(page.locator("#status")).toContainText("2 Links wurden eingefügt.");

  await page.setInputFiles("#file", { name: "Nachtrag.pdf", mimeType: "application/pdf", buffer: Buffer.alloc(2000, 0x41) });
  await page.click("#submit"); // click 2: another file
  await expect.poll(() => calls.filter((c) => c.path.endsWith("/complete")).length).toBe(3);
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");

  const bodies = calls.filter((c) => c.path.endsWith("/complete")).map((c) => JSON.parse(c.body!.toString()));
  expect(bodies[0].step).toBe(bodies[1].step); // same click → same password and password mail
  expect(bodies[0].password).toBe(bodies[1].password);
  expect(bodies[2].step).not.toBe(bodies[0].step); // next click → new password, own password mail
  expect(bodies[2].password).not.toBe(bodies[0].password);
  expect(bodies[2].jobId).toBe("job-1"); // still one mail
  await expectAccessible(page);
});

test("Kennwort an eine andere E-Mail-Adresse: stored on the job once it exists", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await expect(page.locator("#password-to")).toBeHidden();
  await page.check("#channel-mail");
  await expect(page.locator("#password-to-hint")).toContainText("Nur diese Adresse erhält das Kennwort");

  await page.fill("#password-to", "keine adresse");
  await page.press("#password-to", "Tab");
  await expect(page.locator("#status")).toContainText("gültige E-Mail-Adresse");

  await page.fill("#password-to", "kollege@example.com");
  await page.press("#password-to", "Tab");
  await expect(page.locator("#status")).toContainText("nur an kollege@example.com");
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"]).passwordTo).toBe("kollege@example.com");
  await expectAccessible(page);

  // No job yet – after the first link the address goes to the new job.
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect.poll(() => calls.filter((c) => c.method === "PATCH").map((c) => JSON.parse(c.body!.toString())))
    .toEqual([{ passwordTo: "kollege@example.com" }]);
});

test("Kennwort per SMS: not offered without seven.io on the server", async ({ page }) => {
  await openTaskpane(page);
  await expect(page.locator("#channel-sms-row")).toBeHidden(); // no SMS without seven.io on the server
});

test("Kennwort per SMS: number required, normalized to E.164 and stored on the job", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)], config: { smsEnabled: true, smsDefaultCountry: "49" } });
  await page.check("#channel-sms");
  await expect(page.locator("#password-sms")).toBeVisible();
  await expect(page.locator("#password-to-hint")).toContainText("per SMS an diese Nummer");

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte die Mobilnummer angeben");
  await expect(page.locator("#password-sms")).toBeFocused();

  await page.fill("#password-sms", "12");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#status")).toContainText("Bitte eine gültige Mobilnummer angeben");

  await page.fill("#password-sms", "0170 / 123 45 67");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#password-sms")).toHaveValue("+491701234567");
  await expect(page.locator("#status")).toContainText("per SMS an +491701234567");
  await expectAccessible(page);

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect.poll(() => calls.filter((c) => c.method === "PATCH").map((c) => JSON.parse(c.body!.toString())))
    .toEqual([{ passwordSms: "+491701234567" }]);
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"])).toMatchObject({ passwordSms: "+491701234567" });
});

test("Kennwort an eine andere E-Mail-Adresse chosen without an address: no link until it is there", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await page.check("#channel-mail");

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte die E-Mail-Adresse angeben, an die das Kennwort gehen soll");
  await expect(page.locator("#password-to")).toBeFocused();
  await expect(page.locator("#password-to")).toHaveAttribute("aria-invalid", "true");
  await page.click("#request");
  await expect(page.locator("#status")).toContainText("Bitte die E-Mail-Adresse angeben");
  expect(calls.some((c) => c.path === "/api/uploads" || c.path === "/api/requests")).toBe(false); // nothing shared

  // Typed and clicked right away (no "change" yet): the address still reaches the job.
  await page.fill("#password-to", "kollege@example.com");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect.poll(() => calls.filter((c) => c.method === "PATCH").map((c) => JSON.parse(c.body!.toString())))
    .toContainEqual({ passwordTo: "kollege@example.com" });
});

test("links deleted from the mail again: password options unlock, no password mail", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Bericht.pdf", 1.5 * MiB)] });
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect(page.locator("#mode-none")).toBeDisabled(); // mail has a password link
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"]).links).toEqual(["u1"]);

  // The user deletes the card from the mail; the pane notices on focus.
  await page.evaluate(() => {
    (window as unknown as { __officeLog: { inserted: unknown[] } }).__officeLog.inserted.length = 0;
    window.dispatchEvent(new Event("focus"));
  });
  await expect(page.locator("#status")).toContainText("Links wurden aus der Nachricht entfernt");
  await expect(page.locator("#mode-random")).toBeEnabled();
  await expect(page.locator("#mode-none")).toBeEnabled();
  const log = await officeLog(page);
  expect(log.customProps.cloudshareJob).toBeUndefined(); // the send handler has nothing to arm
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({});
});

test("cursor never in the body (new Outlook): the card goes to the top of the mail", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)], cursorOutsideBody: true });
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const log = await officeLog(page);
  expect(log.inserted).toHaveLength(1);
  expect(log.inserted[0].at).toBe("top");
  expect(log.inserted[0].data).toContain("https://cloud.example.com/s/u1");
});

test("Outlook takes no insertion at all: honest error instead of a false success", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)], bodyRejectsInsert: true });
  await page.click("#submit");
  const status = page.locator("#status");
  await expect(status).toContainText("Outlook hat den Link nicht in die Nachricht übernommen");
  await expect(status).toContainText("Meine Freigaben");
  await expect(status).not.toContainText("Der Link wurde eingefügt");
});

test("plain-text mail gets the links as text", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Daten.zip", 1.5 * MiB)], bodyType: "text" });
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const [inserted] = (await officeLog(page)).inserted;
  expect(inserted.coercionType).toBe("text");
  expect(inserted.data).toContain("https://cloud.example.com/s/u1");
  expect(inserted.data).not.toContain("<table");
});

test("Dateien vom Computer hinzufügen: native file dialog (File System Access API)", async ({ page }) => {
  // Stub of window.showOpenFilePicker: returns two chosen files.
  await page.addInitScript(() => {
    (window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async (options: { multiple: boolean }) => {
      (window as unknown as { __pickerOptions: unknown }).__pickerOptions = options;
      const w = window as unknown as { __pickerCalls?: number };
      w.__pickerCalls = (w.__pickerCalls ?? 0) + 1;
      return [
        { getFile: async () => new File(["A".repeat(3000)], "Plan.pdf", { type: "application/pdf" }) },
        { getFile: async () => new File(["B".repeat(2000)], "Fotos.zip", { type: "application/zip" }) },
      ];
    };
  });
  await openTaskpane(page);
  await page.click("#add-files");
  await expect(page.locator("#picked label.list-item")).toHaveCount(2);
  expect(await page.evaluate(() => (window as unknown as { __pickerOptions: unknown }).__pickerOptions)).toEqual({ multiple: true });
  const calls = () => page.evaluate(() => (window as unknown as { __pickerCalls?: number }).__pickerCalls);
  expect(await calls()).toBe(1); // the button click bubbles to the box – still only one dialog
  await page.click("#dropzone .dropzone-icon"); // the whole box is clickable
  expect(await calls()).toBe(2);
});

test("Dateien vom Computer hinzufügen: falls back to the file input (also when the dialog API fails)", async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => {
      throw new DOMException("not allowed here", "SecurityError");
    };
  });
  await openTaskpane(page);
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.click("#add-files")]);
  expect(chooser.isMultiple()).toBe(true);
  await chooser.setFiles({ name: "Plan.pdf", mimeType: "application/pdf", buffer: Buffer.alloc(3000, 0x41) });
  await expect(page.locator("#picked label.list-item")).toHaveCount(1);
});

test("Dateien vom Computer hinzufügen: closing the native dialog adds nothing and opens no second dialog", async ({ page }) => {
  await page.addInitScript(() => {
    (window as unknown as { showOpenFilePicker: unknown }).showOpenFilePicker = async () => {
      throw new DOMException("The user aborted a request.", "AbortError");
    };
  });
  await openTaskpane(page);
  let choosers = 0;
  page.on("filechooser", () => choosers++);
  await page.click("#add-files");
  await page.waitForTimeout(300);
  expect(choosers).toBe(0);
  await expect(page.locator("#picked label.list-item")).toHaveCount(0);
});

test("files dragged from Explorer onto the pane are added (no file dialog needed)", async ({ page }) => {
  await openTaskpane(page);
  const zone = page.locator("#dropzone");
  await expect(zone.locator(".idle")).toHaveText("Dateien hierher ziehen");
  await expect(zone.locator(".drop")).toBeHidden();
  const dataTransfer = await page.evaluateHandle(() => {
    const transfer = new DataTransfer();
    transfer.items.add(new File(["A".repeat(2000)], "Skizze.pdf", { type: "application/pdf" }));
    transfer.items.add(new File(["B".repeat(3000)], "Fotos.zip", { type: "application/zip" }));
    return transfer;
  });
  await page.dispatchEvent("#panel-send", "dragenter", { dataTransfer });
  await expect(page.locator("#panel-send")).toHaveClass(/dropping/);
  await expect(zone.locator(".drop")).toHaveText("Dateien hier ablegen"); // the box turns blue while dragging
  await expect(zone.locator(".drop")).toBeVisible();
  await page.dispatchEvent("#panel-send", "drop", { dataTransfer });
  await expect(page.locator("#panel-send")).not.toHaveClass(/dropping/);
  await expect(page.locator("#picked label.list-item")).toHaveCount(2);
  await expect(page.locator("#status")).toContainText("2 Dateien wurden hinzugefügt.");
});

test("files picked from the computer can be removed from the list again (×)", async ({ page }) => {
  await openTaskpane(page);
  const dataTransfer = await page.evaluateHandle(() => {
    const transfer = new DataTransfer();
    for (const name of ["Eins.pdf", "Zwei.pdf", "Drei.pdf"]) transfer.items.add(new File(["x".repeat(2000)], name));
    return transfer;
  });
  await page.dispatchEvent("#panel-send", "drop", { dataTransfer });
  const rows = page.locator("#picked label.list-item");
  await expect(rows).toHaveCount(3);

  await rows.nth(2).locator(".file-name").click(); // like a user: click the row → "Drei.pdf" not to be uploaded
  await expect(rows.nth(2).locator("input")).not.toBeChecked();
  await expect(rows.nth(2).locator("input")).toHaveAttribute("aria-label", "Drei.pdf, 2 KB");
  await rows.nth(0).locator(".remove-file").click(); // remove "Eins.pdf"
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0)).toContainText("Zwei.pdf");
  await expect(rows.nth(0).locator("input")).toBeChecked(); // the × did not toggle anything
  await expect(rows.nth(1)).toContainText("Drei.pdf");
  await expect(rows.nth(1).locator("input")).not.toBeChecked(); // stays unticked after re-rendering
  await expect(rows.nth(0).locator(".remove-file")).toHaveAttribute("aria-label", "„Zwei.pdf“ aus der Liste entfernen");
  await expectAccessible(page);

  await rows.nth(0).locator(".remove-file").click();
  await rows.nth(0).locator(".remove-file").click();
  await expect(rows).toHaveCount(0);
  await expect(page.locator("#no-attachments")).toBeVisible();
});

test("uploads a file picked from the computer", async ({ page }) => {
  const calls = await openTaskpane(page);
  await expect(page.locator("#no-attachments")).toBeVisible();
  await page.setInputFiles("#file", { name: "Video.mp4", mimeType: "video/mp4", buffer: Buffer.alloc(1.5 * MiB, 7) });
  await expect(page.locator("#picked label.list-item")).toHaveCount(1);
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(calls.filter((c) => c.method === "PUT")).toHaveLength(2);
  expect((await officeLog(page)).inserted[0].data).toContain("Video.mp4");
  await expect(page.locator("#picked label.list-item")).toHaveCount(0);
});

test("own password: live rule check, weak passwords are not accepted", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Bericht.pdf", 1.5 * MiB)] });
  await page.check("#mode-own");
  const password = page.locator("#password");
  await expect(password).toHaveValue("");
  await expect(password).toHaveAttribute("type", "password");
  await expect(page.locator("#password-rules")).toBeVisible();

  await password.fill("abc");
  await expect(page.locator('#password-rules li[data-rule="length"]')).toHaveClass(/fail/);
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Das Kennwort erfüllt noch nicht alle Anforderungen.");
  expect(calls.some((c) => c.path === "/api/uploads")).toBe(false);

  await password.fill(STRONG);
  for (const rule of ["length", "classes", "variety"]) {
    await expect(page.locator(`#password-rules li[data-rule="${rule}"]`)).toHaveClass(/ok/);
  }
  await page.click("#reveal");
  await expect(password).toHaveAttribute("type", "text");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const complete = calls.find((c) => c.path.endsWith("/complete"))!;
  expect(JSON.parse(complete.body!.toString()).password).toBe(STRONG);
});

test("Ohne Kennwort: plain link, no password mail, upload links stay protected", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Flyer.pdf", 1.5 * MiB)] });
  await page.check("#mode-none");
  await expect(page.locator("#password-group")).toBeHidden();
  await expect(page.locator("#max-access-field")).toBeHidden();
  await expect(page.locator("#password-hint")).toContainText("Jeder, der den Link kennt");
  await expectAccessible(page);

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("ohne Kennwort abrufbar");
  const complete = calls.find((c) => c.path.endsWith("/complete"))!;
  const body = JSON.parse(complete.body!.toString());
  expect(body).toMatchObject({ noPassword: true, expireDays: 14 });
  expect(body.password).toBeUndefined();

  const log = await officeLog(page);
  expect(log.inserted[0].data).toContain("https://cloud.example.com/s/u1");
  expect(log.inserted[0].data).not.toContain("Kennwort");
  expect(log.customProps.cloudshareJob).toBeUndefined(); // nothing for the send handler to arm
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({ noPassword: true, links: ["u1"] });

  // The whole mail stays without password; an upload link would need one.
  await expect(page.locator("#mode-random")).toBeDisabled();
  await expect(page.locator("#mode-none")).toBeChecked();
  await page.click("#request");
  await expect(page.locator("#status")).toContainText("Links zum Hochladen brauchen immer ein Kennwort");
  expect(calls.some((c) => c.path === "/api/requests")).toBe(false);
  await shot(page, "no-password");
});

test("Nextcloud enforces link passwords: no 'Ohne Kennwort' option", async ({ page }) => {
  await openTaskpane(page, { passwordEnforced: true });
  await expect(page.locator("#account")).toContainText("max"); // connection status has arrived
  await expect(page.locator("label:has(#mode-none)")).toBeHidden();
  await expect(page.locator("#mode-random")).toBeChecked();
  await expect(page.locator("#password-group")).toBeVisible();
});

test("Dateien von Empfängern erhalten inserts an upload link with the same password", async ({ page }) => {
  const calls = await openTaskpane(page);
  const password = await page.locator("#password").inputValue();
  await page.click("#request");
  await expect(page.locator("#status")).toContainText("Der Link zum Hochladen wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path === "/api/requests")!.body!.toString()).password).toBe(password);
  expect((await officeLog(page)).inserted[0].data).toContain("https://cloud.example.com/s/request");
});

test("revoked Nextcloud app password: returns to the connect view", async ({ page }) => {
  await openTaskpane(page, {
    attachments: [attachment("a1", "Bericht.pdf", 1.5 * MiB)],
    uploadError: { status: 409, error: "Die Nextcloud-Verbindung ist nicht mehr gültig. Bitte neu verbinden.", code: "NEXTCLOUD_NOT_CONNECTED" },
  });
  await page.click("#submit");
  await expect(page.locator("#connect-view")).toBeVisible();
  await expect(page.locator("#status")).toContainText("nicht mehr gültig");
  expect((await officeLog(page)).removed).toEqual([]); // attachment stays in the mail
});

// --- My shares ---------------------------------------------------------------------------------------------------

const SHARES = [
  { id: "11", kind: "file", name: "Präsentation.pptx", url: "https://cloud.example.com/s/a", sharedAt: Date.parse("2026-11-10"),
    expiration: "2099-11-24", downloads: { count: 2, last: "2026-11-11T09:30:00Z" }, received: null, protected: true,
    password: { channel: "sms", to: "+49 170 •••• 67", status: "delivered", at: Date.parse("2026-11-10T10:00:00Z"), deliveredAt: Date.parse("2026-11-10T10:00:09Z") } },
  { id: "12", kind: "file", name: "Angebot.pdf", url: "https://cloud.example.com/s/b", sharedAt: Date.parse("2026-11-12"),
    expiration: "2099-11-26", downloads: { count: 0, last: null }, received: null, protected: false },
  { id: "13", kind: "request", name: "Upload-Link", url: "https://cloud.example.com/s/c", sharedAt: Date.parse("2026-11-01"),
    expiration: "2099-11-15", downloads: null, received: 3 },
  { id: "14", kind: "file", name: "Alt.zip", url: "https://cloud.example.com/s/d", sharedAt: Date.parse("2020-01-01"),
    expiration: "2020-01-15", downloads: { count: 0, last: null }, received: null },
];

test("Meine Freigaben: shows the state of each link and revokes after confirmation", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await page.click("#tab-shares");
  const rows = page.locator("#shares .share");
  await expect(rows).toHaveCount(4);
  await expect(rows.nth(0).locator(".state-text")).toContainText("2× heruntergeladen");
  await expect(rows.nth(1).locator(".state-text")).toHaveText("Noch nicht heruntergeladen");
  await expect(rows.nth(2).locator(".state-text")).toHaveText("3 Dateien empfangen");
  await expect(rows.nth(3).locator(".state-text")).toHaveText("Abgelaufen");
  await expect(rows.nth(0).locator(".share-meta").first()).toContainText("mit Kennwort");
  await expect(rows.nth(0).locator(".share-password")).toContainText("Kennwort per SMS an +49 170 •••• 67 – zugestellt um");
  await expect(rows.nth(1).locator(".share-meta").first()).toContainText("ohne Kennwort");
  await expect(rows.nth(1).locator(".share-password")).toBeHidden();
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "shares-light");

  const revoke = rows.nth(1).locator(".revoke");
  await revoke.click();
  await expect(revoke).toHaveText("Wirklich?");
  expect(calls.some((c) => c.method === "DELETE")).toBe(false); // first click only asks
  await revoke.click();
  await expect(rows).toHaveCount(3);
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/12"]);
  await expect(page.locator("#status")).toContainText("„Angebot.pdf“ ist nicht mehr abrufbar.");
});

test("Meine Freigaben: Als Link einfügen creates a new link with this mail's password", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  const password = await page.locator("#password").inputValue();
  await page.click("#tab-shares");
  await page.locator("#shares .share").nth(1).locator(".insert").click();

  await expect(page.locator("#status")).toContainText("Ein neuer Link zu „Angebot.pdf“ wurde eingefügt.");
  const link = calls.find((c) => c.path === "/api/shares/12/link")!;
  expect(JSON.parse(link.body!.toString())).toMatchObject({ password, expireDays: 14, maxAccess: 3 });
  const log = await officeLog(page);
  expect(log.inserted[0].data).toContain("https://cloud.example.com/s/new-12");
  expect(log.inserted[0].data).toContain("Angebot.pdf");
  expect(log.inserted[0].data).toContain("Das Kennwort erhalten Sie in einer separaten Nachricht.");
  expect(log.customProps.cloudshareJob).toBe("job-1"); // the password mail goes out with this mail
  expect(calls.some((c) => c.method === "DELETE")).toBe(false); // the old link stays
});

test("Meine Freigaben: upload links are not inserted again without password", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await page.check("#mode-none");
  await page.click("#tab-shares");
  await page.locator("#shares .share").nth(2).locator(".insert").click();
  await expect(page.locator("#status")).toContainText("Links zum Hochladen brauchen immer ein Kennwort");
  expect(calls.some((c) => c.path.endsWith("/link"))).toBe(false);
});

// --- Look ----------------------------------------------------------------------------------------------------------

test("dark Office theme: dark pane, still accessible", async ({ page }) => {
  await openTaskpane(page, { theme: "#1f1f1f", attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)], shares: SHARES });
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(page.locator("#attachments label.list-item")).toHaveCount(1);
  await expectAccessible(page);
  await shot(page, "send-dark");
  await page.click("#tab-shares");
  await expect(page.locator("#shares .share")).toHaveCount(4);
  await expectAccessible(page);
  await shot(page, "shares-dark");
});

test("narrow taskpane (320 px): no horizontal scrolling", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await openTaskpane(page, { attachments: [attachment("a1", "Ein sehr langer Dateiname für die Präsentation Q4 2026.pptx", 2.5 * MiB)] });
  await expect(page.locator("#attachments label.list-item")).toHaveCount(1);
  await expectNoHorizontalScroll(page);
  await shot(page, "send-320px");
});
