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
  /** Server with download confirmation: links "Ohne Kennwort" get a job (no password) – see attachNoticeJob. */
  noticeJob?: boolean;
  /** New Outlook: cursor never in the body → setSelectedDataAsync inserts nothing; bodyRejectsInsert: nothing works. */
  cursorOutsideBody?: boolean;
  bodyRejectsInsert?: boolean;
  /** Office SSO failure (plain object as Office delivers it). */
  tokenError?: { code: number; name: string; message: string };
  /** Error response for POST /api/uploads, e.g. Nextcloud connection revoked. */
  uploadError?: { status: number; error: string; code?: string };
  /** Nextcloud enforces passwords for public links. */
  passwordEnforced?: boolean;
  /** Pane with the English texts (served below /en/ by serve.mjs) – for the README pictures. */
  lang?: "en";
  /** Opened from a mail being read (taskpane.html?view=read) with this HTML body. */
  readBody?: string;
}

interface ApiCall { method: string; path: string; body: Buffer | null }

/**
 * GET /api/config as the server answers it (server/src/app.ts) – every field, with the server's defaults except where
 * a test needs something else: a 1 MB threshold (small fixtures), "enforce" and a footer template.
 */
const CONFIG = {
  thresholdMb: 1,
  attachmentPolicy: "enforce",
  defaultExpireDays: 14,
  maxExpireDays: 90,
  passwordLinkMaxDays: 30, // links with password last only as long as their Vaultwarden Send
  chunkSize: 1 * MiB,
  sendMaxAccessCount: 3,
  sendMaxAccessLimit: 20,
  mailFooter: "{displayName} · {jobTitle}",
  trace: false,
  smsEnabled: false,
  smsDefaultCountry: "49",
  smsCountryCodes: [] as string[], // SMS_ALLOWED_COUNTRIES empty = all countries
  smsCountries: "",
  smsAllowedNumbers: "", // SMS_ALLOWED_NUMBERS empty = every number
  smsBlockedNumbers: "^\\+49(800|900|180)", // default of SMS_BLOCKED_NUMBERS
  messageMaxLength: 1000,
  sendExpireDays: 5,
  sendMaxExpireDays: 30,
};

/** Opens the taskpane with the given mail and backend state; returns the list of API calls made by the pane. */
async function openTaskpane(page: Page, scenario: Scenario = {}): Promise<ApiCall[]> {
  const calls: ApiCall[] = [];
  let connected = scenario.connected ?? true;
  const uploads = new Map<string, { fileName: string; size: number }>();

  await page.addInitScript((cfg) => {
    (window as unknown as { __officeConfig: unknown }).__officeConfig = cfg;
  }, { attachments: scenario.attachments ?? [], theme: scenario.theme, bodyType: scenario.bodyType, tokenError: scenario.tokenError,
    cursorOutsideBody: scenario.cursorOutsideBody, bodyRejectsInsert: scenario.bodyRejectsInsert, readBody: scenario.readBody });
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
      const noPassword = JSON.parse(call.body!.toString()).noPassword === true;
      const jobId = noPassword ? (scenario.noticeJob ? "notice-1" : null) : "job-1";
      return json({ shareUrl: `https://cloud.example.com/s/${id}`, expireDate: "2026-12-01", jobId, noPassword, fileName, size });
    }
    if (route_ === "POST /api/requests") {
      return json({ shareUrl: "https://cloud.example.com/s/request", expireDate: "2026-12-01", jobId: "job-1" });
    }
    if (route_ === "POST /api/messages") {
      return json({ shareUrl: "https://vault.example.com/#/send/msg/key123", expireDate: "2026-12-01", jobId: "job-1" });
    }
    if (route_ === "GET /api/shares") return json(scenario.shares ?? []);
    if (/^DELETE \/api\/shares\/.+$/.test(route_)) return json({ ok: true });
    if (/^PATCH \/api\/jobs\/[^/]+$/.test(route_)) return json({ maxAccess: JSON.parse(call.body!.toString()).maxAccess });
    if (/^GET \/api\/shares\/\d+\/password$/.test(route_)) {
      return json({ password: "KateGeteFi45", sms: "+491701234567", otherTo: [],
        recipients: { to: ["erika@example.com"], cc: ["chef@example.com"], bcc: ["archiv@example.com"] } });
    }
    if (/^POST \/api\/shares\/\d+\/password\/send$/.test(route_)) {
      const { channel, to } = JSON.parse(call.body!.toString());
      return json({ channel, to: channel === "sms" ? "+49 171 •••• 21" : channel === "recipients" ? "erika@example.com, chef@example.com" : to });
    }
    if (/^POST \/api\/shares\/\d+\/link$/.test(route_)) {
      const id = call.path.split("/")[3];
      const share = (scenario.shares as Array<{ id: string; kind: string; name: string }> | undefined)?.find((s) => s.id === id);
      const jobId = JSON.parse(call.body!.toString()).noPassword === true ? null : "job-1";
      return json({ kind: share?.kind ?? "file", shareUrl: `https://cloud.example.com/s/new-${id}`, expireDate: "2026-12-01",
        jobId, fileName: share?.name ?? "x", size: share?.kind === "request" ? null : 3 * MiB });
    }
    return json({ error: `not mocked: ${route_}` }, 404);
  });

  const view = scenario.readBody === undefined ? "" : "?view=read";
  await page.goto(scenario.lang ? `/${scenario.lang}/taskpane.html${view}` : `/taskpane.html${view}`);
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

/** "Dateien" / "Nachricht" / "Anfordern" – click the visible segment like a user (the radio is styled away). */
const chooseKind = (page: Page, kind: "files" | "message" | "request") => page.click(`label:has(#kind-${kind})`);

/** Opens the collapsed "Gültigkeit" section (lifetimes, accesses). */
async function openValidity(page: Page) {
  await page.click("#validity-toggle");
  await expect(page.locator("#validity-toggle")).toHaveAttribute("aria-expanded", "true");
}

const shot = async (page: Page, name: string) =>
  test.info().attach(name, { body: await page.screenshot({ fullPage: true }), contentType: "image/png" });

/**
 * Screenshot for the README: the viewport grows until the scrolling content fits, so the sticky footer covers nothing
 * (the pane pictures in README.md come from these attachments).
 */
async function fullShot(page: Page, name: string) {
  const size = page.viewportSize()!;
  // The document scrolls and the footer sticks to the bottom of the viewport: make the viewport as tall as the
  // document, then take a plain (not full-page) screenshot.
  const height = await page.evaluate(() => document.documentElement.scrollHeight);
  await page.setViewportSize({ width: size.width, height: Math.max(size.height, height) });
  await page.waitForTimeout(100); // layout after the resize
  test.info().attach(name, { body: await page.screenshot(), contentType: "image/png" });
  await page.setViewportSize(size);
}

/**
 * Picture for the README (attachment "readme-<name>"): exactly as tall as the pane's content – the viewport is first
 * made small so the full-height layout collapses, then set to the content height (sticky footer at its natural place).
 */
async function readmeShot(page: Page, name: string) {
  const size = page.viewportSize()!;
  await page.setViewportSize({ width: size.width, height: 200 });
  await page.waitForTimeout(100);
  const height = await page.evaluate(() => document.documentElement.scrollHeight);
  await page.setViewportSize({ width: size.width, height });
  await page.waitForTimeout(100); // layout after the resize
  test.info().attach(`readme-${name}`, { body: await page.screenshot(), contentType: "image/png" });
  await page.setViewportSize(size);
}

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
    config: { smsEnabled: true, smsDefaultCountry: "49" }, // the README picture shows every way the password can go
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
  // One button that says what it inserts; the rarely changed values collapsed behind a summary.
  await expect(page.locator("#kind-files")).toBeChecked();
  await expect(page.locator("#submit")).toHaveText("Datei als Link einfügen");
  await rows.nth(1).locator(".file-name").click();
  await expect(page.locator("#submit")).toHaveText("2 Dateien als Link einfügen");
  await rows.nth(1).locator(".file-name").click();
  await expect(page.locator("#validity-section")).toBeHidden();
  await expect(page.locator("#validity-summary")).toHaveText("Link 14 Tage · Kennwort-Link 5 Tage, 3-mal abrufbar");
  await expect(page.locator('#password-channel option[value="sms"]')).toHaveCount(1);

  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await fullShot(page, "send-light");
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
  await openValidity(page);
  await page.fill("#max-access", "7");
  await page.press("#max-access", "Tab");
  await expect(page.locator("#status")).toContainText("7-mal abrufbar (für den nächsten Link)");
  expect(calls.some((c) => c.method === "PATCH")).toBe(false);
});

test("Link gültig für / Kennwort abrufbar: capped at the maximum with a note, confirmed otherwise", async ({ page }) => {
  await openTaskpane(page);
  const status = page.locator("#status");
  await openValidity(page);
  await expect(page.locator("#validity-summary")).toBeHidden(); // the fields themselves are visible now
  // Every field explains itself – "Link gültig für" by what the link is for.
  await expect(page.locator("#expire-hint")).toHaveText("So lange können die Empfänger die Dateien herunterladen – danach ist der Link gesperrt.");
  await chooseKind(page, "request");
  await expect(page.locator("#expire-hint")).toHaveText("So lange können die Empfänger über den Link Dateien hochladen – danach ist er gesperrt.");
  await chooseKind(page, "files");

  // With a password a link lasts at most as long as its password in Vaultwarden: 30 days.
  await page.fill("#expire", "500");
  await page.press("#expire", "Tab");
  await expect(page.locator("#expire")).toHaveValue("30");
  await expect(status).toContainText("Höchstens 30 Tage möglich – der Wert wurde auf 30 gesetzt.");
  await expect(status).toContainText("Neue Links sind 30 Tage gültig.");

  // Without password up to SHARE_MAX_EXPIRE_DAYS; back with a password it is capped again.
  await page.click("label:has(#mode-none)");
  await page.fill("#expire", "500");
  await page.press("#expire", "Tab");
  await expect(page.locator("#expire")).toHaveValue("90");
  await expect(status).toContainText("Höchstens 90 Tage möglich – der Wert wurde auf 90 gesetzt.");
  await page.click("label:has(#mode-random)");
  await expect(page.locator("#expire")).toHaveValue("30");

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

  // Closed again: the summary shows the new values.
  await page.click("#validity-toggle");
  await expect(page.locator("#validity-summary")).toHaveText("Link 10 Tage · Kennwort-Link 5 Tage, einmal abrufbar");
});

test("every 'Als Link einfügen' gets its own password; files of one click share it", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Angebot.pdf", 1.5 * MiB), attachment("a2", "Plan.zip", 1.5 * MiB)],
  });
  await expect(page.locator("#submit")).toHaveText("2 Dateien als Link einfügen");
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
  await page.selectOption("#password-channel", "mail");
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
  await expect(page.locator("#password-channel option")).toHaveText(["Empfänger der Nachricht", "Andere E-Mail-Adresse"]);
  await expect(page.locator('#password-channel option[value="sms"]')).toHaveCount(0); // no SMS without seven.io on the server
});

test("Kennwort per SMS: number required, normalized to E.164 and stored on the job", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)], config: { smsEnabled: true, smsDefaultCountry: "49" } });
  await page.selectOption("#password-channel", "sms");
  await expect(page.locator("#password-sms")).toBeVisible();
  await expect(page.locator("#password-to-hint")).toContainText("in einer SMS an diese Nummer");

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte die Rufnummer angeben");
  await expect(page.locator("#password-sms")).toBeFocused();

  await page.fill("#password-sms", "12");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#status")).toContainText("Bitte eine gültige Rufnummer angeben");

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

test("Kennwort per SMS: numbers of other countries and service numbers are refused right away", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)],
    config: { smsEnabled: true, smsCountryCodes: ["49", "43", "41"], smsCountries: "Deutschland, Österreich und Schweiz" },
  });
  await page.selectOption("#password-channel", "sms");
  await expect(page.locator("#password-to-hint")).toContainText("Möglich sind Rufnummern aus Deutschland, Österreich und Schweiz.");

  await page.fill("#password-sms", "+44 7700 900123");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#status")).toContainText("nur an Rufnummern aus Deutschland, Österreich und Schweiz");
  await expect(page.locator("#password-sms")).toHaveAttribute("aria-invalid", "true");

  await page.click("#submit"); // no link with a number the password could never reach
  await expect(page.locator("#status")).toContainText("nur an Rufnummern aus Deutschland, Österreich und Schweiz");
  expect(calls.filter((c) => c.method === "PATCH")).toEqual([]);

  await page.fill("#password-sms", "0800 123 45 67"); // service number (SMS_BLOCKED_NUMBERS)
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#status")).toContainText("Service- und Sondernummern wie 0800, 0180 oder 0900 sind ausgeschlossen");
  await expect(page.locator("#password-sms")).toHaveAttribute("aria-invalid", "true");

  await page.fill("#password-sms", "+43 (0)664 123 45 67");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#password-sms")).toHaveValue("+436641234567");
  await expect(page.locator("#password-sms")).not.toHaveAttribute("aria-invalid", "true");
  await expect(page.locator("#status")).toContainText("per SMS an +436641234567");
});

test("Kennwort an eine andere E-Mail-Adresse chosen without an address: no link until it is there", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await page.selectOption("#password-channel", "mail");

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte die E-Mail-Adresse angeben, an die das Kennwort gehen soll");
  await expect(page.locator("#password-to")).toBeFocused();
  await expect(page.locator("#password-to")).toHaveAttribute("aria-invalid", "true");
  await chooseKind(page, "request");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte die E-Mail-Adresse angeben");
  expect(calls.some((c) => c.path === "/api/uploads" || c.path === "/api/requests")).toBe(false); // nothing shared

  // Typed and clicked right away (no "change" yet): the address still reaches the job.
  await chooseKind(page, "files");
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
  // Dragged in while "Nachricht" is shown: the pane switches to "Dateien" so the drop box is there.
  await chooseKind(page, "message");
  await expect(zone).toBeHidden();
  await page.dispatchEvent("#panel-send", "dragenter", { dataTransfer });
  await expect(page.locator("#kind-files")).toBeChecked();
  await expect(page.locator("#panel-send")).toHaveClass(/dropping/);
  await expect(zone.locator(".drop")).toHaveText("Dateien hier ablegen"); // the box turns blue while dragging
  await expect(zone.locator(".drop")).toBeVisible();
  await page.dispatchEvent("#panel-send", "drop", { dataTransfer });
  await expect(page.locator("#panel-send")).not.toHaveClass(/dropping/);
  await expect(page.locator("#picked label.list-item")).toHaveCount(2);
  await expect(page.locator("#status")).toContainText("2 Dateien wurden hinzugefügt.");
  await expect(page.locator("#submit")).toHaveText("2 Dateien als Link einfügen");
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

  // The whole mail stays without password; an upload link would need one – said before any click.
  await expect(page.locator("#mode-random")).toBeDisabled();
  await expect(page.locator("#mode-none")).toBeChecked();
  await expect(page.locator("#validity-summary")).toHaveText("Link 14 Tage");
  await chooseKind(page, "request");
  await expect(page.locator("#password-hint")).toContainText("Links zum Hochladen brauchen ein Kennwort");
  await expect(page.locator("#submit")).toBeDisabled();
  await chooseKind(page, "message");
  await expect(page.locator("#password-hint")).toContainText("Eine vertrauliche Nachricht braucht ein Kennwort");
  await expect(page.locator("#submit")).toBeDisabled();
  expect(calls.some((c) => c.path === "/api/requests" || c.path === "/api/messages")).toBe(false);
  await expectAccessible(page);
  await shot(page, "no-password");
  await chooseKind(page, "files");
  await expect(page.locator("#submit")).toBeEnabled();
});

test("Nachricht / Anfordern: 'Ohne Kennwort' is not offered – a choice made for files becomes a random password", async ({ page }) => {
  await openTaskpane(page);
  await page.click("label:has(#mode-none)");
  await expect(page.locator("#password-group")).toBeHidden();
  await chooseKind(page, "request");
  await expect(page.locator("#mode-none-row")).toBeHidden();
  await expect(page.locator("#mode-random")).toBeChecked();
  await expect(page.locator("#password-group")).toBeVisible();
  await expect(page.locator("#password")).not.toHaveValue("");
  await expect(page.locator("#submit")).toHaveText("Link zum Hochladen einfügen");
  await chooseKind(page, "files");
  await expect(page.locator("#mode-none-row")).toBeVisible();
  await expect(page.locator("#submit")).toHaveText("Als Link einfügen"); // no attachments, nothing ticked
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
  await chooseKind(page, "request");
  await expect(page.locator("#request-section")).toBeVisible();
  await expect(page.locator("#files-section")).toBeHidden();
  await expectAccessible(page);
  await shot(page, "request");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link zum Hochladen wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path === "/api/requests")!.body!.toString()).password).toBe(password);
  expect((await officeLog(page)).inserted[0].data).toContain("https://cloud.example.com/s/request");
});

test("Nachricht: inserts a card with the link, never the text, and shares the password step", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)],
    config: { messageMaxLength: 1000, sendExpireDays: 5, sendMaxExpireDays: 30 },
  });
  await expect(page.locator("#message-section")).toBeHidden();
  await chooseKind(page, "message");
  await expect(page.locator("#message-section")).toBeVisible();
  await expect(page.locator("#files-section")).toBeHidden(); // one thing at a time
  await expect(page.locator("#submit")).toHaveText("Vertrauliche Nachricht einfügen"); // the only button
  await expect(page.locator("#mode-none-row")).toBeHidden(); // a confidential message always has a password
  await expect(page.locator("#validity-summary")).toHaveText("Nachricht 5 Tage · Kennwort-Link 5 Tage, 3-mal abrufbar");

  const password = await page.locator("#password").inputValue();
  await page.fill("#message", "Zugang zum Portal: anna / Sommer2026");
  await expect(page.locator("#message-count")).toHaveText("36 von 1000 Zeichen");
  await expectAccessible(page);
  await fullShot(page, "message-light");
  await openValidity(page);
  await expect(page.locator("#expire-field")).toBeHidden(); // no Nextcloud link for a text
  await expect(page.locator("#message-days-hint")).toHaveText("So lange können die Empfänger die Nachricht mit dem Kennwort öffnen – danach wird sie gelöscht.");
  await expect(page.locator("#message-days")).toHaveValue("5");
  await page.fill("#message-days", "45");
  await page.locator("#message-days").blur();
  await expect(page.locator("#message-days")).toHaveValue("30");
  await expect(page.locator("#status")).toContainText("Höchstens 30 Tage möglich");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link zur vertraulichen Nachricht wurde eingefügt.");

  const body = JSON.parse(calls.find((c) => c.path === "/api/messages")!.body!.toString());
  expect(body).toMatchObject({ text: "Zugang zum Portal: anna / Sommer2026", password, messageDays: 30, passwordDays: 5 });
  const inserted = (await officeLog(page)).inserted[0].data;
  expect(inserted).toContain("https://vault.example.com/#/send/msg/key123");
  expect(inserted).toContain("Vertrauliche Nachricht für Sie");
  expect(inserted).not.toContain("Sommer2026"); // the text never goes into the mail
  await expect(page.locator("#message")).toHaveValue("");
  expect(calls.some((c) => c.path === "/api/uploads")).toBe(false); // only the text, no files
  expect((await officeLog(page)).removed).toEqual([]); // the attachment was not touched
  await expectAccessible(page);

  // Empty text: says so instead of inserting anything.
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Bitte zuerst die vertrauliche Nachricht eingeben.");
  await expect(page.locator("#message")).toBeFocused();
  expect(calls.filter((c) => c.path === "/api/messages")).toHaveLength(1);
});

test("Kennwort-Link gültig für: 5 days by default, at most 30, sent with every link", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await openValidity(page);
  await expect(page.locator("#password-days")).toHaveValue("5");
  await page.fill("#password-days", "40");
  await page.locator("#password-days").blur(); // leaving the field fires "change" (like a user)
  await expect(page.locator("#password-days")).toHaveValue("30");
  await expect(page.locator("#status")).toContainText("Höchstens 30 Tage möglich – der Wert wurde auf 30 gesetzt. Neue Kennwort-Links sind 30 Tage gültig.");
  // "Ohne Kennwort" has no password link (click the visible label – the radio itself is styled away).
  await page.click("label:has(#mode-none)");
  await expect(page.locator("#mode-none")).toBeChecked();
  await expect(page.locator("#password-days-field")).toBeHidden();
  await page.click("label:has(#mode-random)");
  await expect(page.locator("#password-days-field")).toBeVisible();

  await page.fill("#password-days", "12");
  await page.locator("#password-days").blur();
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString()).passwordDays).toBe(12);
});

test("Kennwort per SMS: no password-link settings – the password is in the SMS and works as long as the link", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)], config: { smsEnabled: true } });
  await openValidity(page);
  await expect(page.locator("#max-access-field")).toBeVisible();
  await expect(page.locator("#password-days-field")).toBeVisible();
  await page.selectOption("#password-channel", "sms");
  await expect(page.locator("#max-access-field")).toBeHidden();
  await expect(page.locator("#password-days-field")).toBeHidden();
  await expect(page.locator("#password-to-hint")).toContainText("gilt, solange der Link gültig ist");
  await page.fill("#password-sms", "0170 1234567");
  await page.locator("#password-sms").blur();
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  // The password link only holds the password until sending – as long as possible.
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString()).passwordDays).toBe(30);

  await page.selectOption("#password-channel", "recipients");
  await expect(page.locator("#password-days-field")).toBeVisible();
});

test("Ohne Kennwort with download confirmation: the draft is tagged so the confirmation knows the sent mail", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Flyer.pdf", 1.5 * MiB)], noticeJob: true });
  await page.click("label:has(#mode-none)");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("ohne Kennwort abrufbar");
  const log = await officeLog(page);
  expect(log.customProps.cloudshareJob).toBe("notice-1");
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({ noPassword: true, jobId: "notice-1", links: ["u1"] });
  expect(log.inserted[0].data).not.toContain("Kennwort");
  await expect(page.locator("#mode-random")).toBeDisabled();
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
  { id: "12", kind: "file", name: "Angebot.pdf", url: "https://cloud.example.com/s/b", sharedAt: Date.parse("2099-11-12"),
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
  // All valid links, no expired one ("Alt.zip"), no refresh button and no "last 30 days".
  await expect(rows).toHaveCount(3);
  await expect(page.locator("#shares")).not.toContainText("Alt.zip");
  await expect(page.locator("#refresh-shares")).toHaveCount(0);
  await expect(rows.nth(0).locator(".state-text")).toContainText("2× heruntergeladen");
  await expect(rows.nth(1).locator(".state-text")).toHaveText("Noch nicht heruntergeladen");
  await expect(rows.nth(2).locator(".state-text")).toHaveText("3 Dateien empfangen");
  // One short line: period and password – the word "Kennwort" only once.
  const meta = rows.nth(0).locator(".share-meta").first();
  await expect(meta.locator(".share-period")).toHaveText("10.11.2026 – 24.11.2099");
  await expect(meta.locator(".share-password")).toContainText("Kennwort per SMS an +49 170 •••• 67 – zugestellt um");
  expect((await meta.textContent())!.match(/Kennwort/g)).toHaveLength(1);
  await expect(rows.nth(1).locator(".share-period")).toHaveText("12.11. – 26.11.2099"); // same year: no year at the start
  await expect(rows.nth(1).locator(".share-password")).toHaveText("ohne Kennwort");
  await expect(rows.nth(2).locator(".share-password")).toHaveText("mit Kennwort");
  // Icons instead of words; the words are the buttons' names.
  await expect(rows.nth(0).locator(".copy")).toHaveAttribute("aria-label", "Link kopieren");
  await expect(rows.nth(0).locator(".insert")).toHaveAttribute("aria-label", "Als neuen Link in diese Nachricht einfügen");
  await expect(rows.nth(0).locator(".revoke")).toHaveAttribute("aria-label", "Freigabe zurückziehen");
  await expect(rows.nth(0).locator(".share-actions")).toHaveText("");
  await rows.nth(0).locator(".copy").click();
  await expect(rows.nth(0).locator(".copy")).toHaveAttribute("aria-label", "Kopiert");
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "shares-light");

  const revoke = rows.nth(1).locator(".revoke");
  await revoke.click();
  await expect(revoke).toHaveText("Wirklich?");
  expect(calls.some((c) => c.method === "DELETE")).toBe(false); // first click only asks
  await revoke.click();
  await expect(rows).toHaveCount(2);
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/12"]);
  await expect(page.locator("#status")).toContainText("„Angebot.pdf“ ist nicht mehr abrufbar.");
});

test("Meine Freigaben: a failed SMS says whether it was sent at all, and how to send the password again", async ({ page }) => {
  const sms = { channel: "sms", to: "+49 170 •••• 67", status: "failed", at: Date.parse("2026-11-10T10:00:00Z"), deliveredAt: null };
  const share = (id: string, name: string, refused: boolean) => ({ ...SHARES[0], id, name, url: `https://cloud.example.com/s/${id}`,
    password: { ...sms, refused }, passwordAvailable: true });
  await openTaskpane(page, { shares: [share("21", "Refused.pdf", true), share("22", "Lost.pdf", false)] });
  await page.click("#tab-shares");
  const rows = page.locator("#shares .share");
  // seven.io refused it right away (key, credit, limit).
  await expect(rows.nth(0).locator(".share-password")).toHaveText("Kennwort per SMS an +49 170 •••• 67 – nicht versendet – bitte erneut senden");
  // seven.io accepted it and reported "not delivered" later.
  await expect(rows.nth(1).locator(".share-password")).toHaveText("Kennwort per SMS an +49 170 •••• 67 – nicht zugestellt – bitte erneut senden");
  await expect(rows.nth(1).locator(".share-password")).toHaveClass(/failed/);
});

test("Meine Freigaben: links of a mail not sent (e.g. a discarded draft) say so and when they are revoked", async ({ page }) => {
  const revokeAt = Date.parse("2099-11-22T10:00:00Z");
  await openTaskpane(page, {
    shares: [
      { ...SHARES[0], id: "31", password: null, unsent: true, revokeAt },
      { ...SHARES[0], id: "32", password: null, unsent: true, revokeAt: null },
      SHARES[1],
    ],
  });
  await page.click("#tab-shares");
  const rows = page.locator("#shares .share");
  await expect(rows.nth(0).locator(".share-unsent")).toHaveText(
    `Nicht gesendet – wird die Nachricht bis 22.11.2099 nicht gesendet, zieht CloudShare den Link zurück.`,
  );
  await expect(rows.nth(1).locator(".share-unsent")).toHaveText("Nicht gesendet – das Kennwort wurde noch nicht verschickt.");
  await expect(rows.nth(2).locator(".share-unsent")).toBeHidden();
  await expectAccessible(page);
});

test("Meine Freigaben: Kennwort shows the password and sends it again the same ways as the first time", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: [{ ...SHARES[0], passwordAvailable: true }, SHARES[1]], config: { smsEnabled: true } });
  await page.click("#tab-shares");
  const rows = page.locator("#shares .share");
  await expect(rows.nth(1).locator(".show-password")).toBeHidden(); // no password to show (link without password)
  const toggle = rows.nth(0).locator(".show-password");
  await expect(toggle).toHaveAttribute("aria-label", "Kennwort von „Präsentation.pptx“ anzeigen und erneut senden");
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  const panel = rows.nth(0).locator(".password-panel");
  await expect(panel.locator("input.mono")).toHaveValue("KateGeteFi45");
  // The password went by SMS: SMS to the same number is suggested.
  await expect(panel.locator("select")).toHaveValue("sms");
  await expect(panel.locator(".channel-input")).toHaveValue("+491701234567");
  await expectAccessible(page);
  await shot(page, "password-panel");

  // The same choices as when sending: recipients of the message, another address, SMS.
  await expect(panel.locator("select option")).toHaveText(["Empfänger der Nachricht", "Andere E-Mail-Adresse", "Rufnummer (SMS)"]);

  // Recipients of the message: To, Cc and Bcc as the first time – nothing to type, the hint names them.
  await panel.locator("select").selectOption("recipients");
  await expect(panel.locator(".channel-input")).toBeHidden();
  await expect(panel.locator(".hint")).toHaveText("Wie beim ersten Mal – An: erika@example.com · Cc: chef@example.com · Bcc: archiv@example.com");
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toHaveText(
    "Das Kennwort wurde wie beim ersten Mal an die Empfänger der Nachricht geschickt (erika@example.com, chef@example.com).");
  const sends = () => calls.filter((c) => c.path === "/api/shares/11/password/send").map((c) => JSON.parse(c.body!.toString()));
  expect(sends()[0]).toEqual({ channel: "recipients", to: "" });

  // Other addresses: each gets a mail of its own (server), typed here.
  // The list reloads after sending (new state of the password) – wait for it, then open the panel again.
  await expect.poll(() => calls.filter((c) => c.path === "/api/shares").length).toBe(2);
  await expect(panel).toBeHidden();
  await toggle.click();
  await panel.locator("select").selectOption("mail");
  await expect(panel.locator(".channel-input")).toHaveValue("");
  await expect(panel.locator(".hint")).toHaveText("Jede Adresse erhält eine eigene Nachricht mit einem Link zum Kennwort.");
  await panel.locator(".channel-input").fill("kollege@example.com, buergerin@example.com");
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toHaveText("Das Kennwort wurde an kollege@example.com, buergerin@example.com geschickt.");
  expect(sends()[1]).toEqual({ channel: "mail", to: "kollege@example.com, buergerin@example.com" });
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

/** A sent mail with the cards of "Präsentation.pptx" (/s/a) and the upload link (/s/c), plus an unrelated link. */
const SENT_MAIL = '<p>Hallo,</p><a href="https://cloud.example.com/s/a">Herunterladen</a>' +
  '<a href="https://cloud.example.com/s/c">Dateien hochladen</a><a href="https://example.org/s/zzz">Shop</a>';

test("reading a mail: its links on top, the other shares below, nothing to insert", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES, readBody: SENT_MAIL });
  await expect(page.locator("#tab-send")).toBeHidden(); // no "Senden" tab when reading
  await expect(page.locator("#tab-shares")).toHaveAttribute("aria-selected", "true");
  await expect(page.locator("#help-toggle")).toBeVisible(); // the guide is there too
  await expect(page.locator("#panel-send")).toBeHidden();
  const inMail = page.locator("#mail-shares .share");
  await expect(inMail).toHaveCount(2);
  await expect(inMail.nth(0).locator(".share-name")).toHaveText("Präsentation.pptx");
  await expect(inMail.nth(0).locator(".share-password")).toContainText("zugestellt um");
  await expect(inMail.nth(1).locator(".share-name")).toHaveText("Upload-Link");
  await expect(page.locator("#shares .share")).toHaveCount(1); // Angebot.pdf (Alt.zip has expired)
  await expect(page.locator("#mail-none")).toBeHidden();
  await expect(page.locator(".share .insert").first()).toBeHidden(); // a mail being read cannot take a link
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "read-view");
  expect(calls.filter((c) => c.path === "/api/shares")).toHaveLength(1);

  // Pinned pane, the user selects another mail without CloudShare links.
  await page.evaluate(() => {
    (window as unknown as { __officeConfig: { readBody: string } }).__officeConfig.readBody = "<p>Keine Links</p>";
    (window as unknown as { __officeHandlers: Record<string, () => void> }).__officeHandlers.olkItemSelectedChanged();
  });
  await expect(page.locator("#mail-none")).toHaveText("Diese Nachricht enthält keine CloudShare-Links von Ihnen.");
  await expect(page.locator("#mail-shares .share")).toHaveCount(0);
  await expect(page.locator("#shares .share")).toHaveCount(3);

  // Revoking works here too.
  const revoke = page.locator("#shares .share").nth(1).locator(".revoke");
  await revoke.click();
  await revoke.click();
  await expect(page.locator("#shares .share")).toHaveCount(2);
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/12"]);
});

test("reading a mail before connecting Nextcloud: connect view, then the shares", async ({ page }) => {
  await openTaskpane(page, { connected: false, shares: SHARES, readBody: SENT_MAIL });
  await expect(page.locator("#connect-view")).toBeVisible();
  await page.click("#connect");
  await expect(page.locator("#mail-shares .share")).toHaveCount(2, { timeout: 10_000 });
  await expect(page.locator("#tab-send")).toBeHidden();
});

test("Meine Freigaben: upload links are not inserted again without password", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await page.check("#mode-none");
  await page.click("#tab-shares");
  await page.locator("#shares .share").nth(2).locator(".insert").click();
  await expect(page.locator("#status")).toContainText("Links zum Hochladen brauchen immer ein Kennwort");
  expect(calls.some((c) => c.path.endsWith("/link"))).toBe(false);
});

// --- Help -----------------------------------------------------------------------------------------------------------

test("Hilfe: the \"?\" at the top right opens the guide in place of the tab, readable at 320 px", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await openTaskpane(page);
  await expectNoHorizontalScroll(page); // tabs and "?" fit
  const toggle = page.locator("#help-toggle");
  await expect(toggle).toHaveAttribute("aria-label", "Hilfe");
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  await expect(page.locator("#panel-help")).toBeVisible();
  await expect(page.locator("#help-close")).toBeFocused(); // screen readers start in the guide
  await expect(page.locator("#help-heading")).toHaveText("Hilfe");
  await expect(page.locator("#panel-send")).toBeHidden();
  const help = page.locator("#panel-help");
  await expect(help).not.toContainText("help."); // every text exists in the language file
  const sections = help.locator("details");
  await expect(sections).toHaveCount(7);
  // The most common task is open, the rest closed until needed.
  await expect(sections.nth(1).locator("summary")).toHaveText("Große Dateien senden");
  await expect(sections.nth(1)).toHaveAttribute("open", "");
  await expect(sections.nth(6)).not.toHaveAttribute("open", "");
  await sections.nth(6).locator("summary").click();
  await expect(sections.nth(6)).toHaveAttribute("open", "");
  await expect(sections.nth(6)).toContainText("Ich habe den Link an die falsche Person geschickt.");
  await expect(help).toContainText("Personenbezogene oder vertrauliche Daten versenden Sie bitte immer mit Kennwort.");
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "help");

  // "×" closes it: back where the user was, focus on the "?".
  await page.click("#help-close");
  await expect(page.locator("#panel-help")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
  await expect(toggle).toHaveAttribute("aria-expanded", "false");
  await expect(toggle).toBeFocused();

  // Opened from "Meine Freigaben": a click on a tab closes it as well.
  await page.click("#tab-shares");
  await toggle.click();
  await expect(page.locator("#panel-shares")).toBeHidden();
  await toggle.click(); // the "?" again closes it
  await expect(page.locator("#panel-shares")).toBeVisible();
  await toggle.click();
  await page.click("#tab-send");
  await expect(page.locator("#panel-help")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
});

// --- Look ----------------------------------------------------------------------------------------------------------

test("dark Office theme: dark pane, still accessible", async ({ page }) => {
  await openTaskpane(page, { theme: "#1f1f1f", attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)], shares: SHARES, config: { smsEnabled: true } });
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(page.locator("#attachments label.list-item")).toHaveCount(1);
  await expectAccessible(page);
  await fullShot(page, "send-dark");
  await page.click("#tab-shares");
  await expect(page.locator("#shares .share")).toHaveCount(3); // the expired "Alt.zip" is not listed
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

// --- README pictures (English) -------------------------------------------------------------------------------------
// The README shows the pane in English: same mocks, texts from locales/en.json (attachments "readme-*").

test.describe("README pictures (English)", () => {
  test.use({ locale: "en-GB" });

  const ATTACHMENTS = [attachment("a1", "Presentation.pptx", 6.5 * MiB), attachment("a2", "Notes.txt", 200 * 1024)];
  const README_CONFIG = { smsEnabled: true, smsDefaultCountry: "49", thresholdMb: 5, sendExpireDays: 5, sendMaxExpireDays: 30 };
  const SHARES_EN = SHARES.map((share, i) => ({ ...share, name: ["Presentation.pptx", "Quote.pdf", "File request", "Old.zip"][i] }));

  test("Send: files", async ({ page }) => {
    await openTaskpane(page, { lang: "en", attachments: ATTACHMENTS, config: README_CONFIG });
    await expect(page.locator("#submit")).toHaveText("Insert file as link");
    await expect(page.locator("#validity-summary")).toHaveText("Link 14 days · Password link 5 days, opens 3 times");
    await expect(page.locator("#account")).toHaveText("max [12.5 GB of 50 GB free]");
    await expectAccessible(page);
    await readmeShot(page, "send");
  });

  test("Send: confidential message and validity", async ({ page }) => {
    await openTaskpane(page, { lang: "en", attachments: ATTACHMENTS, config: README_CONFIG });
    await chooseKind(page, "message");
    await page.fill("#message", "Portal access: anna / Summer2026");
    await openValidity(page);
    await expect(page.locator("#submit")).toHaveText("Insert confidential message");
    await expectAccessible(page);
    await readmeShot(page, "message");
  });

  test("Send: dark theme", async ({ page }) => {
    await openTaskpane(page, { lang: "en", theme: "#1f1f1f", attachments: ATTACHMENTS, config: README_CONFIG });
    await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
    await expect(page.locator("#attachments label.list-item")).toHaveCount(2);
    await expectAccessible(page);
    await readmeShot(page, "dark");
  });

  test("My shares", async ({ page }) => {
    const shares = SHARES_EN.map((share, i) => (i === 0 ? { ...share, passwordAvailable: true } : share));
    await openTaskpane(page, { lang: "en", shares, config: README_CONFIG });
    await page.click("#tab-shares");
    await expect(page.locator("#shares .share")).toHaveCount(3);
    await expect(page.locator("#shares .share").nth(1).locator(".state-text")).toHaveText("Not downloaded yet");
    // "Password" opened at the first link: the password and "Send again to".
    await page.locator("#shares .share").nth(0).locator(".show-password").click();
    await expect(page.locator("#shares .share").nth(0).locator(".password-panel input.mono")).toHaveValue("KateGeteFi45");
    await expectAccessible(page);
    await readmeShot(page, "shares");
  });

  test("Reading a sent mail", async ({ page }) => {
    await openTaskpane(page, { lang: "en", shares: SHARES_EN, config: README_CONFIG, readBody: SENT_MAIL });
    await expect(page.locator("#mail-shares .share")).toHaveCount(2);
    await expect(page.locator("#mail-heading")).toHaveText("In this message");
    await readmeShot(page, "read");
  });

  test("Help", async ({ page }) => {
    await openTaskpane(page, { lang: "en", config: README_CONFIG });
    await page.click("#help-toggle");
    await expect(page.locator("#panel-help details").nth(1)).toHaveAttribute("open", "");
    await expect(page.locator("#panel-help")).not.toContainText("help.");
    await expectAccessible(page);
    await readmeShot(page, "help");
  });

  test("Connect", async ({ page }) => {
    await openTaskpane(page, { lang: "en", connected: false });
    await expect(page.locator("#connect-view")).toBeVisible();
    await readmeShot(page, "connect");
  });
});
