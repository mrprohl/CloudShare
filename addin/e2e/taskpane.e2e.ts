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
  /** Office SSO failure (plain object as Office delivers it). */
  tokenError?: { code: number; name: string; message: string };
  /** Error response for POST /api/uploads, e.g. Nextcloud connection revoked. */
  uploadError?: { status: number; error: string; code?: string };
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
  }, { attachments: scenario.attachments ?? [], theme: scenario.theme, bodyType: scenario.bodyType, tokenError: scenario.tokenError });
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
      return json(connected ? { connected: true, loginName: "max", freeBytes: 12 * 1024 ** 3 } : { connected: false });
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
      return json({ shareUrl: `https://cloud.example.com/s/${id}`, expireDate: "2026-12-01", jobId: "job-1", fileName, size });
    }
    if (route_ === "POST /api/requests") {
      return json({ shareUrl: "https://cloud.example.com/s/request", expireDate: "2026-12-01", jobId: "job-1" });
    }
    if (route_ === "GET /api/shares") return json(scenario.shares ?? []);
    if (/^DELETE \/api\/shares\/.+$/.test(route_)) return json({ ok: true });
    return json({ error: `not mocked: ${route_}` }, 404);
  });

  await page.goto("/taskpane.html");
  return calls;
}

const officeLog = (page: Page) =>
  page.evaluate(() => (window as unknown as { __officeLog: {
    inserted: { data: string; coercionType: string }[];
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
  expect((await password.inputValue()).length).toBe(20);
  await expect(page.locator("#max-access")).toHaveValue("3");
  await expect(page.locator("#expire")).toHaveValue("14");
  await expect(page.locator("#opt-out-row")).toBeHidden(); // enforce: no way around CloudShare

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
  expect(log.inserted[0].data).toContain("https://cloud.example.com/s/u1");
  expect(log.inserted[0].data).toContain("Präsentation.pptx");
  expect(log.inserted[0].data).toContain("Max Mustermann · Projektleitung"); // footer from the profile
  expect(log.customProps.cloudshareJob).toBe("job-1"); // links the draft to the password mail
  expect(JSON.parse(log.session["cloudshare-state"])).toMatchObject({ jobId: "job-1", password, maxAccess: 3 });

  // From now on the password applies to the whole mail.
  await expect(page.locator("#mode-random")).toBeDisabled();
  await expect(page.locator("#mode-own")).toBeDisabled();
  await expect(page.locator("#max-access")).toHaveAttribute("readonly", "");
  await shot(page, "after-upload");
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

test("Dateien anfordern inserts an upload link with the same password", async ({ page }) => {
  const calls = await openTaskpane(page);
  const password = await page.locator("#password").inputValue();
  await page.click("#request");
  await expect(page.locator("#status")).toContainText("Der Upload-Link wurde eingefügt.");
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
    expiration: "2099-11-24", downloads: { count: 2, last: "2026-11-11T09:30:00Z" }, received: null },
  { id: "12", kind: "file", name: "Angebot.pdf", url: "https://cloud.example.com/s/b", sharedAt: Date.parse("2026-11-12"),
    expiration: "2099-11-26", downloads: { count: 0, last: null }, received: null },
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
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "shares-light");

  const revoke = rows.nth(1).locator(".revoke");
  await revoke.click();
  await expect(revoke).toHaveText("Wirklich zurückziehen?");
  expect(calls.some((c) => c.method === "DELETE")).toBe(false); // first click only asks
  await revoke.click();
  await expect(rows).toHaveCount(3);
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/12"]);
  await expect(page.locator("#status")).toContainText("„Angebot.pdf“ ist nicht mehr abrufbar.");
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
