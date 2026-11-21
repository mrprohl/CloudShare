/**
 * UI tests of the taskpane (built add-in in Chromium at taskpane width). Office.js is replaced by e2e/office-mock.js,
 * the backend by route mocks below – the test checks what the user sees and what the pane does to the mail.
 */
import AxeBuilder from "@axe-core/playwright";
import { expect, test, type Page, type Route } from "@playwright/test";
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
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
  /** Nextcloud login name in the status (default "max.mustermann"). */
  loginName?: string;
  /** Error response for POST /api/uploads, e.g. Nextcloud connection revoked. */
  uploadError?: { status: number; error: string; code?: string };
  /** Nextcloud enforces passwords for public links. */
  passwordEnforced?: boolean;
  /** Pane with the English texts (served below /en/ by serve.mjs) – for the README pictures. */
  lang?: "en";
  /** Display language of the Outlook client (office-mock.js; default: German, below /en/ English). */
  displayLanguage?: string;
  /** Personal settings stored on the server (GET /api/settings), on top of DEFAULT_SETTINGS. */
  settings?: Partial<typeof DEFAULT_SETTINGS>;
  /** Sensitivity labels published to the user (Outlook's catalog) and the draft's own label (GUID). */
  labels?: Array<{ id: string; name: string }>;
  /** The labels /api/labels returns (Graph, InformationProtectionPolicy.Read); unset: null (not consented). */
  tenantLabels?: Array<{ id: string; name: string; rank: number }>;
  /** The own Outlook contacts /api/contacts matches (Contacts.Read); with `type` = a number (SMS); unset: null (not consented). */
  contacts?: Array<{ name: string; value: string; type?: "mobile" | "business" | "home" }>;
  /** More of GET /api/shares/:id/password – e.g. `{ deferred: true, verified: [] }` for a password held back for the proof. */
  passwordInfo?: Record<string, unknown>;
  /** GET /api/proofs: the user's delivery proofs. */
  proofs?: unknown[];
  /** Lists and settings answer only after this many ms – to see the loading placeholders. */
  slowLoads?: number;
  mailLabel?: string;
  /** Opened from a mail being read (taskpane.html?view=read) with this HTML body. */
  readBody?: string;
}

interface ApiCall { method: string; path: string; body: Buffer | null; language?: string }

/**
 * GET /api/config as the server answers it (server/src/app.ts) – every field, with the server's defaults except where
 * a test needs something else: a 1 MB threshold (small fixtures), "enforce" and a footer template.
 */
const CONFIG = {
  // Success messages stay in the tests (one test checks that they run down) – screenshots and checks see them.
  successMessageSeconds: 0,
  nextcloudUrl: "https://cloud.example.com",
  thresholdMb: 1,
  attachmentPolicy: "enforce",
  defaultExpireDays: 14,
  maxExpireDays: 90,
  passwordLinkMaxDays: 30, // links with password last only as long as their Vaultwarden Send
  chunkSize: 1 * MiB,
  sendMaxAccessCount: 3,
  sendMaxAccessLimit: 20,
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
  deliveryProof: false, // DELIVERY_PROOF: the box "Zustellnachweis" is offered only when the server has it
  labelPolicy: "inherit", // LABEL_POLICY default
  labelEncrypted: "warn", // LABEL_ENCRYPTED default
};

/** Personal settings as the server starts every user (DEFAULT_* unset). */
const DEFAULT_SETTINGS = {
  downloadNotice: true,
  uploadNotice: true,
  smsNotice: true,
  channels: ["recipients", "mail", "sms"] as string[],
  defaultChannel: "recipients",
  language: "auto",
  recipientLanguage: "de",
  signature: null as null | Record<string, unknown>,
};

/** Opens the taskpane with the given mail and backend state; returns the list of API calls made by the pane. */
async function openTaskpane(page: Page, scenario: Scenario = {}): Promise<ApiCall[]> {
  const calls: ApiCall[] = [];
  let connected = scenario.connected ?? true;
  // The English README pictures show English as the language for the recipients.
  let userSettings = { ...DEFAULT_SETTINGS, ...(scenario.lang === "en" ? { recipientLanguage: "en" } : {}), ...scenario.settings };
  const uploads = new Map<string, { fileName: string; size: number }>();

  await page.addInitScript((cfg) => {
    (window as unknown as { __officeConfig: unknown }).__officeConfig = cfg;
  }, { attachments: scenario.attachments ?? [], theme: scenario.theme, bodyType: scenario.bodyType, tokenError: scenario.tokenError,
    cursorOutsideBody: scenario.cursorOutsideBody, bodyRejectsInsert: scenario.bodyRejectsInsert, readBody: scenario.readBody,
    displayLanguage: scenario.displayLanguage, labels: scenario.labels, mailLabel: scenario.mailLabel });
  await page.route("https://appsforoffice.microsoft.com/**", (route) =>
    route.fulfill({ contentType: "text/javascript", body: officeMock }),
  );

  await page.route("**/api/**", async (route: Route) => {
    const request = route.request();
    const url = new URL(request.url());
    const call = { method: request.method(), path: url.pathname, body: request.postDataBuffer(),
      language: request.headers()["x-cloudshare-language"] };
    calls.push(call);
    const json = (body: unknown, status = 200) => route.fulfill({ status, contentType: "application/json", body: JSON.stringify(body) });
    const route_ = `${call.method} ${call.path}`;
    if (scenario.slowLoads && /^GET \/api\/(shares|proofs|settings)$/.test(route_)) {
      await new Promise((resolve) => setTimeout(resolve, scenario.slowLoads));
    }

    // "Zustellnachweise": the user's cases, and the signed address of a PDF.
    if (route_ === "GET /api/proofs") return json(scenario.proofs ?? []);

    if (/^POST \/api\/proofs\/[^/]+\/link$/.test(route_)) {
      return json({ url: `https://cloudshare.example.com/proof/${call.path.split("/")[3]}.pdf?until=1&lang=de&sig=x` });
    }
    // The tenant's sensitivity labels (Graph) – null like a server without the permission.
    if (route_ === "GET /api/labels") return json({ labels: scenario.tenantLabels ?? null });
    // "Feedback senden": the server mails it (FEEDBACK_MAIL).
    if (route_ === "POST /api/feedback") return json({ ok: true });
    // Suggestions from the own contacts: numbers for kind=sms, addresses otherwise, matched like the server (roughly).
    if (route_ === "GET /api/contacts") {
      if (!scenario.contacts) return json({ contacts: null });
      const q = (url.searchParams.get("q") ?? "").toLowerCase();
      const sms = url.searchParams.get("kind") === "sms";
      return json({ contacts: scenario.contacts.filter((c) => Boolean(c.type) === sms
        && (c.name.toLowerCase().includes(q) || c.value.toLowerCase().includes(q))) });
    }
    if (route_ === "GET /api/config") {
      return scenario.config === "down" ? json({ error: "down" }, 503) : json({ ...CONFIG, ...scenario.config });
    }
    if (route_ === "GET /api/profile") {
      return json({ displayName: "Max Mustermann", givenName: "Max", surname: "Mustermann", jobTitle: "Projektleitung",
        department: "", company: "Beispiel GmbH", email: "max.mustermann@example.com", phone: "", mobile: "", office: "" });
    }
    if (route_ === "GET /api/nextcloud/status") {
      return json(connected
        ? { connected: true, loginName: scenario.loginName ?? "max.mustermann", freeBytes: 12.5 * 1024 ** 3, usedBytes: 37.5 * 1024 ** 3, totalBytes: 50 * 1024 ** 3,
            passwordEnforced: scenario.passwordEnforced ?? false }
        : { connected: false });
    }
    if (route_ === "DELETE /api/nextcloud/connection") {
      connected = false;
      return json({ ok: true });
    }
    if (route_ === "POST /api/nextcloud/connect") {
      return json({ connectId: "c1", loginUrl: "https://cloud.example.com/login/v2/flow/abc" });
    }
    if (route_ === "POST /api/nextcloud/connect/c1/poll") {
      connected = true;
      return json({ connected: true, loginName: "max.mustermann" });
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
      const request = JSON.parse(call.body!.toString());
      const noPassword = request.noPassword === true;
      const jobId = noPassword ? (scenario.noticeJob ? "notice-1" : null) : "job-1";
      // "Zustellnachweis" ticked: the gateway address and the tracking ID of the case.
      const proof = request.proof === true;
      return json({
        shareUrl: proof ? `https://cloudshare.example.com/r/s/${id}` : `https://cloud.example.com/s/${id}`,
        expireDate: "2026-12-01", jobId, noPassword, fileName, size, proofId: proof ? "20261117-1146-7Q2F" : null,
      });
    }
    if (route_ === "POST /api/requests") {
      // "Zustellnachweis" ticked: the server answers with the gateway address and the tracking ID of the case.
      const proof = JSON.parse(call.body!.toString()).proof === true;
      return json({
        shareUrl: proof ? "https://cloudshare.example.com/r/s/request" : "https://cloud.example.com/s/request",
        expireDate: "2026-12-01", jobId: "job-1", proofId: proof ? "20261117-1146-7Q2F" : null,
      });
    }
    if (route_ === "POST /api/messages") {
      return json({ shareUrl: "https://vault.example.com/#/send/msg/key123", expireDate: "2026-12-01", jobId: "job-1" });
    }
    if (/^(GET|PUT|DELETE) \/api\/settings$/.test(route_)) {
      // Like server/src/settings.ts: SMS only with seven.io, the notices the server has switched on.
      const sms = scenario.config !== "down" && scenario.config?.smsEnabled === true;
      const available = { downloadNotice: true, uploadNotice: true, smsNotice: sms, channels: sms ? ["recipients", "mail", "sms"] : ["recipients", "mail"],
        languages: ["de", "en"] };
      const within = (s: typeof DEFAULT_SETTINGS) => {
        const channels = s.channels.filter((c) => available.channels.includes(c));
        return { ...s, smsNotice: s.smsNotice && sms, channels, defaultChannel: channels.includes(s.defaultChannel) ? s.defaultChannel : channels[0] };
      };
      if (call.method === "PUT") userSettings = { ...userSettings, ...JSON.parse(call.body!.toString()) };
      if (call.method === "DELETE") userSettings = { ...DEFAULT_SETTINGS };
      return json({ settings: within(userSettings), defaults: within(DEFAULT_SETTINGS), available });
    }
    if (route_ === "GET /api/shares") return json(scenario.shares ?? []);
    if (/^DELETE \/api\/shares\/.+$/.test(route_)) return json({ ok: true });
    // "Verlauf": the sent mail of an entry, as the EWS ID Outlook opens.
    if (/^GET \/api\/proofs\/[^/]+\/mail$/.test(route_)) return json({ ewsId: `ews-proof-${route_.split("/")[3]}` });
    if (/^GET \/api\/shares\/[^/]+\/mail$/.test(route_)) return json({ ewsId: `ews-${route_.split("/")[3]}` });
    if (/^PATCH \/api\/jobs\/[^/]+$/.test(route_)) return json({ maxAccess: JSON.parse(call.body!.toString()).maxAccess });
    if (/^GET \/api\/shares\/[^/]+\/password$/.test(route_)) {
      return json({ password: "KateGeteFi45", sms: "+491701234567", otherTo: [],
        recipients: { to: ["erika@example.com"], cc: ["chef@example.com"], bcc: ["archiv@example.com"] }, ...scenario.passwordInfo });
    }
    if (/^POST \/api\/shares\/[^/]+\/password\/send$/.test(route_)) {
      const { channel, to } = JSON.parse(call.body!.toString());
      return json({ channel, to: channel === "sms" ? "+49 171 7654321" : channel === "recipients" ? "erika@example.com, chef@example.com" : to });
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
    inserted: { data: string; coercionType: string; at: "cursor" | "top" | "body" }[];
    removed: string[];
    customProps: Record<string, string>;
    session: Record<string, string>;
    openedUrls: string[];
    openedMails: string[];
    mailLabel?: string;
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
/** "Kennwort an": a radio button – click its visible label (the radio itself is styled away). */
const channelRadio = (page: Page, channel: "recipients" | "mail" | "sms") => page.locator(`#password-channel input[value="${channel}"]`);
const chooseChannel = (page: Page, channel: "recipients" | "mail" | "sms") =>
  page.click(`#password-channel label:has(input[value="${channel}"])`);

/** Opens the guide: account at the bottom left → "Hilfe". */
async function openHelp(page: Page) {
  await page.locator("#account-toggle").click();
  await page.locator("#help-toggle").click();
  await expect(page.locator("#panel-help")).toBeVisible();
}

/** Opens "Verlauf": account at the bottom left → "Verlauf" (a page like the settings). */
async function openHistory(page: Page) {
  await page.locator("#account-toggle").click();
  await page.locator("#shares-open").click();
  await expect(page.locator("#panel-shares")).toBeVisible();
}

/** Opens the personal settings: account at the bottom left → "Einstellungen". */
async function openSettings(page: Page) {
  await page.locator("#account-toggle").click();
  await page.locator("#settings-open").click();
  await expect(page.locator("#panel-settings")).toBeVisible();
}

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

/** Fresh README pictures of this run (artifact "readme-pictures") – copied to docs/images when the pane changed. */
const README_PICTURES = path.join(path.dirname(fileURLToPath(import.meta.url)), "../readme-pictures");

/**
 * Picture for the README (attachment "readme-<name>"): exactly as tall as the pane's content – the viewport is first
 * made small so the full-height layout collapses, then set to the content height (sticky footer at its natural place).
 * It must match docs/images/taskpane-<name>.png: the documentation shows the pane as it is (README → "Pictures").
 */
async function readmeShot(page: Page, name: string) {
  const size = page.viewportSize()!;
  await page.setViewportSize({ width: size.width, height: 200 });
  await page.waitForTimeout(100);
  const height = await page.evaluate(() => document.documentElement.scrollHeight);
  await page.setViewportSize({ width: size.width, height });
  await page.waitForTimeout(100); // layout after the resize
  // A generated password differs on every run: the picture shows a fixed example, so it can be compared.
  await page.evaluate(() => {
    const field = document.querySelector<HTMLInputElement>("#password");
    if (field && field.readOnly && field.value) field.value = "VusMoveSira64";
  });
  const picture = await page.screenshot({ animations: "disabled", caret: "hide" });
  test.info().attach(`readme-${name}`, { body: picture, contentType: "image/png" });
  mkdirSync(README_PICTURES, { recursive: true });
  writeFileSync(path.join(README_PICTURES, `taskpane-${name}.png`), picture);
  await page.setViewportSize(size);
  expect(
    picture,
    `docs/images/taskpane-${name}.png no longer shows the pane – take it from the artifact "readme-pictures" (README → "Pictures")`,
  ).toMatchSnapshot(`taskpane-${name}.png`, { maxDiffPixels: 0 }); // CI renders the same pixels every time; a changed word must fail
}

// --- Connection --------------------------------------------------------------------------------------------------

test("a progress message is shown whole (its class must not clash with the progress bar)", async ({ page }) => {
  await openTaskpane(page);
  // The kind of the message is its class – "progress" once also styled the 2px progress bar and cut the message.
  const height = await page.evaluate(() => {
    const bar = document.getElementById("status")!;
    bar.hidden = false;
    bar.className = "messagebar progress";
    bar.innerHTML = '<svg class="icon" viewBox="0 0 16 16"></svg><span>„cloud-redirect.zip“ wird hochgeladen … 42 %</span>';
    return bar.getBoundingClientRect().height;
  });
  expect(height).toBeGreaterThanOrEqual(36);
});

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
  await expect(page.locator("#account")).toContainText("max.mustermann");
  expect(calls.some((c) => c.path === "/api/nextcloud/connect/c1/poll")).toBe(true);
});

test("connecting works with the message timer on (an empty message has no ring)", async ({ page }) => {
  // The timer is off in the other tests; clearing the bar (setStatus("")) once looked for a ring that was not there.
  await openTaskpane(page, { connected: false, config: { successMessageSeconds: 5 } });
  await page.click("#connect");
  await expect(page.locator("#send-view")).toBeVisible({ timeout: 10_000 });
  await expect(page.locator("#status")).toContainText("Die Verbindung zu Nextcloud wurde hergestellt.");
  await expect(page.locator("#status")).not.toContainText("fehlgeschlagen");
});

test("opening the pane: settings and connection status share one sign-in (no 13008)", async ({ page }) => {
  // Settings and Nextcloud status both need a token right at startup.
  const calls = await openTaskpane(page);
  await expect(page.locator("#send-view")).toBeVisible();
  await expect.poll(() => calls.some((c) => c.path === "/api/settings")).toBe(true);
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
  // "Erneut versuchen" starts CloudShare afresh (the page loads again and asks Office once more).
  const retry = status.locator("button.status-action");
  await expect(retry).toHaveText("Erneut versuchen");
  await expectAccessible(page);
  await Promise.all([page.waitForEvent("load"), retry.click()]);
  await expect(page.locator("#status")).toContainText("(Code 13007)");
  await shot(page, "sso-error");
});

test("server unreachable: says so instead of showing an empty pane", async ({ page }) => {
  await openTaskpane(page, { config: "down" });
  await expect(page.locator("#status")).toContainText("CloudShare ist derzeit nicht erreichbar");
  await expect(page.locator("#status button.status-action")).toHaveText("Erneut versuchen");
});

// --- Language of the Outlook client --------------------------------------------------------------------------------

test("language: the pane follows Outlook's display language and tells the server", async ({ page }) => {
  // The page comes in German (LANGUAGE), the client is English: the English page replaces it before the pane starts.
  const calls = await openTaskpane(page, { displayLanguage: "en-US" });
  await expect(page.locator("#send-view")).toBeVisible();
  await expect(page.locator("html")).toHaveAttribute("lang", "en");
  await expect(page.locator("label:has(#kind-files) span")).toHaveText("Files");
  await expect(page.locator("#shares-open")).toHaveText("History");
  await expect(page.locator("#password-label")).toHaveText("Password");
  await expect.poll(() => calls.filter((c) => c.path.startsWith("/api/")).length).toBeGreaterThan(0);
  expect(calls.every((c) => c.language === "en")).toBe(true); // error messages of the server in English
});

test("language: a client language without texts falls back to English", async ({ page }) => {
  const calls = await openTaskpane(page, { displayLanguage: "fr-FR" });
  await expect(page.locator("#send-view")).toBeVisible();
  await expect(page.locator("label:has(#kind-files) span")).toHaveText("Files");
  expect(calls.every((c) => c.language === "en")).toBe(true);
});

// --- Personal settings (gear in the account line) -------------------------------------------------------------------

test("settings: notices and signature in sight, the rest under \"Weitere Einstellungen\"; every change saved at once", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Plan.pdf", 2 * MiB)], config: { smsEnabled: true } });
  await expect(page.locator("#send-view")).toBeVisible();
  await openSettings(page);
  await expect(page.locator("#panel-settings")).toBeVisible();
  await expect(page.locator("#panel-send")).toBeHidden();
  await expect(page.locator("#settings-form")).toBeVisible();
  // All ways offered, notices on, language as in Outlook, the organisation's signature.
  await expect(page.locator('#settings-channels input[type="checkbox"]')).toHaveCount(3);
  await expect(page.locator("#settings-notices input")).toHaveCount(3);
  await expect(page.locator("#settings-language")).toHaveValue("auto");
  // Under the heading "Sprache" each select has its own label, like "Sprache für die Empfänger".
  await expect(page.locator('label[for="settings-language"]')).toHaveText("Sprache von CloudShare");
  await expect(page.locator("#settings-signature-format")).toHaveValue("none");
  // In sight: notices and signature; folded away: the ways for the password, the languages and "Zurücksetzen".
  await expect(page.locator("#settings-notices-group")).toBeVisible();
  await expect(page.locator("#settings-signature-format")).toBeVisible();
  await expect(page.locator("#settings-default-channel")).toBeHidden();
  await expect(page.locator("#settings-language")).toBeHidden();
  await expect(page.locator("#settings-reset")).toBeHidden();
  await expect(page.locator("#settings-save")).toHaveCount(0); // no "Speichern" any more
  await expectAccessible(page);
  await shot(page, "settings");

  // A notice: saved with the click.
  // The real checkbox is invisible under its box (as in the attachment list): click the row like a user.
  const row = (input: string) => page.locator("label.checkbox").filter({ has: page.locator(input) });
  const puts = () => calls.filter((c) => c.method === "PUT" && c.path === "/api/settings").map((c) => JSON.parse(c.body!.toString()));
  await row('input[name="downloadNotice"]').click();
  await expect(page.locator('#settings-notices input[name="downloadNotice"]')).not.toBeChecked();
  await expect(page.locator("#status")).toHaveText("Ihre Einstellungen wurden gespeichert.");
  await expect.poll(() => puts().length).toBe(1);
  expect(puts()[0]).toMatchObject({ downloadNotice: false, uploadNotice: true, smsNotice: true, language: "auto", signature: null });

  // Only "Andere E-Mail-Adresse" and SMS, SMS preselected – under "Weitere Einstellungen".
  await page.locator("#settings-more > summary").click();
  await expect(page.locator("#settings-default-channel")).toBeVisible();
  await expect(page.locator("#settings-reset")).toBeVisible();
  await expectAccessible(page);
  // SMS preselected first, then "Empfänger der Nachricht" left out: each saved at once, the mail (no links yet) moves
  // to the preselected way as soon as its own is no longer offered.
  await page.locator("#settings-default-channel").selectOption("sms");
  await row('input[value="recipients"]').click();
  await expect(page.locator('#settings-channels input[value="recipients"]')).not.toBeChecked();
  await expect.poll(() => puts().at(-1)).toMatchObject({ channels: ["mail", "sms"], defaultChannel: "sms", downloadNotice: false });

  // The settings stay open; back to "Senden" with the arrow.
  await expect(page.locator("#panel-settings")).toBeVisible();
  await page.click("#settings-close");
  await expect(page.locator("#panel-send")).toBeVisible();
  await expect(page.locator("#account-toggle")).toBeFocused();
  // No links in this mail yet: it switches to the preselected way, the ways left out are not offered.
  await expect(channelRadio(page, "sms")).toBeChecked();
  await expect(page.locator('#password-channel label:has(input[value="recipients"])')).toBeHidden();
  await expect(page.locator('#password-channel label:has(input[value="mail"])')).toBeVisible();
});

test("settings and help: the heading with its back arrow stays in sight while scrolling", async ({ page }) => {
  await page.setViewportSize({ width: 340, height: 520 });
  await openTaskpane(page, { config: { smsEnabled: true } });
  await openSettings(page);
  await expect(page.locator("#settings-form")).toBeVisible();
  await page.locator("#settings-more").scrollIntoViewIfNeeded(); // scrolled to the bottom
  await expect(page.locator("#settings-close")).toBeInViewport();
  await expect(page.locator("#settings-heading")).toBeInViewport();
  await page.click("#settings-close");
  await expect(page.locator("#panel-settings")).toBeHidden();

  await openHelp(page);
  await page.locator(".help-support").scrollIntoViewIfNeeded();
  await expect(page.locator("#help-close")).toBeInViewport();
});

test("settings: a new mail starts with the preselected way and offers only the chosen ones", async ({ page }) => {
  await openTaskpane(page, { config: { smsEnabled: true }, settings: { channels: ["recipients", "sms"], defaultChannel: "sms" } });
  await expect(channelRadio(page, "sms")).toBeChecked();
  await expect(page.locator("#password-sms")).toBeVisible();
  await expect(page.locator('#password-channel label:has(input[value="mail"])')).toBeHidden();
  await expect(page.locator('#password-channel label:has(input[value="recipients"])')).toBeVisible();
});

test("settings: an own signature is plain text with paragraphs – stored as typed, not in the link card", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Plan.pdf", 2 * MiB)] });
  await openSettings(page);
  await expect(page.locator("#settings-signature-format option")).toHaveText(["Vorgabe der Organisation", "Eigene Signatur"]);
  await page.locator("#settings-signature-format").selectOption("text");
  const text = page.locator("#settings-signature-text");
  await expect(text).toBeVisible();
  await expect(page.locator("#settings-signature-hint")).toHaveText("Steht als Text am Ende jeder E-Mail, die CloudShare an Ihre Empfänger sendet – Kennwort und Bestätigungscode.");
  await text.fill("Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone}");
  // The "i" folds out the placeholders; a click inserts one at the cursor.
  const info = page.locator("#settings-placeholders-toggle");
  await expect(page.locator("#settings-placeholders")).toBeHidden();
  await info.click();
  await expect(info).toHaveAttribute("aria-expanded", "true");
  const list = page.locator("#settings-placeholder-list li");
  await expect(list).toHaveCount(10);
  await expect(list.filter({ hasText: "Mobiltelefon" }).locator("button")).toHaveText("{mobile}");
  await expectAccessible(page);
  await text.press("End"); // cursor at the end of the text
  await text.pressSequentially(" · Mobil ");
  await list.filter({ hasText: "Mobiltelefon" }).locator("button").click();
  await expect(text).toHaveValue("Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone} · Mobil {mobile}");
  await info.click();
  await expect(page.locator("#settings-placeholders")).toBeHidden();
  await text.fill("Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone}");
  // Saved a moment after typing stops – no button.
  await expect(page.locator("#status")).toHaveText("Ihre Einstellungen wurden gespeichert.");
  await expect.poll(() => {
    const puts = calls.filter((c) => c.method === "PUT");
    return puts.length ? JSON.parse(puts.at(-1)!.body!.toString()).signature : null;
  }).toEqual({ format: "text", content: "Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone}" });

  // Opened again: the text as typed.
  await page.click("#settings-close");
  await openSettings(page);
  await expect(page.locator("#settings-signature-format")).toHaveValue("text");
  await expect(text).toHaveValue("Mit freundlichen Grüßen\n\n{displayName}\nTel. {phone}");

  // The link card carries no signature – it ends the password mail.
  await page.click("#settings-close");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const inserted = (await officeLog(page)).inserted[0].data as string;
  expect(inserted).not.toContain("Mit freundlichen Grüßen");
});

test("settings: a language chosen in the settings beats Outlook's", async ({ page }) => {
  const calls = await openTaskpane(page, { settings: { language: "en" } });
  // Loaded once, cached in roamingSettings, the pane starts again in English.
  await expect(page.locator("html")).toHaveAttribute("lang", "en");
  await expect(page.locator("label:has(#kind-files) span")).toHaveText("Files");
  await expect.poll(() => calls.filter((c) => c.path === "/api/config").length).toBeGreaterThan(1);
  await openSettings(page);
  await expect(page.locator("#settings-language")).toHaveValue("en");
  await expect(page.locator("#settings-language option")).toHaveText(["Default (as in Outlook)", "Deutsch", "English"]);
});

test("language: a regional variant uses its language", async ({ page }) => {
  const calls = await openTaskpane(page, { displayLanguage: "de-AT" });
  await expect(page.locator("#send-view")).toBeVisible();
  await expect(page.locator("html")).toHaveAttribute("lang", "de");
  await expect(page.locator("label:has(#kind-files) span")).toHaveText("Dateien");
  expect(calls.every((c) => c.language === "de")).toBe(true);
});

// --- Sending files -----------------------------------------------------------------------------------------------

test("ticks every attachment and proposes a random password", async ({ page }) => {
  await openTaskpane(page, {
    attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB), attachment("a2", "Notiz.txt", 200 * 1024)],
    config: { smsEnabled: true, smsDefaultCountry: "49" }, // the README picture shows every way the password can go
  });
  const rows = page.locator("#attachments label.list-item");
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0).locator("input")).toBeChecked();
  // Whoever opens CloudShare sends through it: every attachment ticked, also the small one; no "Groß" mark, an "×" each.
  await expect(page.locator("#attachments .badge")).toHaveCount(0);
  await expect(rows.nth(0).locator(".remove-file")).toHaveAttribute("aria-label", "„Präsentation.pptx“ entfernen");
  await expect(rows.nth(1).locator("input")).toBeChecked();

  const password = page.locator("#password");
  await expect(password).toHaveAttribute("readonly", "");
  expect(await password.inputValue()).toMatch(/^(?=(?:[^2-9]*[2-9]){2}[^2-9]*$)[A-Za-z2-9]{12,15}$/); // pronounceable, two digits
  await expect(page.locator("#max-access")).toHaveValue("3");
  await expect(page.locator("#expire")).toHaveValue("14");
  await expect(page.locator("#account")).toHaveText("max.mustermann");
  await expect(page.locator("#account-storage")).toHaveText("37,5 GB von 50 GB belegt"); // in the account menu
  // One button that says what it inserts; the rarely changed values collapsed behind a summary.
  await expect(page.locator("#kind-files")).toBeChecked();
  await expect(page.locator("#submit")).toHaveText("2 Dateien als Link einfügen");
  await rows.nth(1).locator(".file-name").click();
  await expect(page.locator("#submit")).toHaveText("Datei als Link einfügen");
  await rows.nth(1).locator(".file-name").click();
  await expect(page.locator("#validity-section")).toBeHidden();
  await expect(page.locator("#validity-summary")).toHaveText("Link 14 Tage · Kennwort-Link 5 Tage, 3-mal abrufbar");
  await expect(channelRadio(page, "sms")).toHaveCount(1);

  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await fullShot(page, "send-light");
});

test("a subject typed after inserting the link reaches \"Verlauf\" before sending", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)] });
  await page.evaluate(() => {
    (window as unknown as { __officeConfig: { subject: string } }).__officeConfig.subject = ""; // no subject yet
  });
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString())).toMatchObject({ subject: "" });

  // Typed afterwards: the next check (every few seconds) tells the server; unchanged it asks nothing.
  await page.evaluate(() => {
    (window as unknown as { __officeConfig: { subject: string } }).__officeConfig.subject = "Unterlagen Neubau";
  });
  const patches = () => calls.filter((c) => c.method === "PATCH" && c.path.startsWith("/api/jobs/"));
  await expect.poll(() => patches().length, { timeout: 10_000 }).toBe(1);
  expect(JSON.parse(patches()[0].body!.toString())).toEqual({ subject: "Unterlagen Neubau" });
  await page.waitForTimeout(4500);
  expect(patches()).toHaveLength(1);
});

test("Sprache für die Empfänger: only from the settings – the card and the password mail in it, the pane stays German", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)],
    settings: { recipientLanguage: "en" },
  });
  await expect(page.locator("#password")).not.toHaveValue("");
  // No choice per mail in "Senden": the language is set under "Einstellungen" → "Weitere Einstellungen" only.
  await expect(page.locator("#recipient-language")).toHaveCount(0);
  await expect(page.locator("#validity-summary")).not.toContainText("English");
  await page.click("#submit");

  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt."); // the pane in its own language
  const card = (await officeLog(page)).inserted[0].data as string;
  expect(card).toContain("A file has been shared with you");
  expect(card).toContain(">Download<");
  expect(card).not.toContain("Herunterladen");
  // The server writes the password mail and SMS of this step in it.
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString())).toMatchObject({ language: "en", subject: "Unterlagen Neubau" }); // the draft's subject for "Verlauf"
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"] ?? "{}")).not.toHaveProperty("recipientLanguage");
});

test("the \"×\" of an attachment removes it from the mail", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB), attachment("a2", "Notiz.txt", 200 * 1024)] });
  const rows = page.locator("#attachments label.list-item");
  await expect(rows).toHaveCount(2);
  await rows.nth(1).locator(".remove-file").click();
  expect((await officeLog(page)).removed).toEqual(["a2"]);
  await expect(rows).toHaveCount(1);
  await expect(rows.nth(0).locator(".file-name")).toHaveText("Präsentation.pptx");
});

test("uploads a large attachment in chunks, removes it and inserts the card", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)] });
  await expect(page.locator("#password")).not.toHaveValue(""); // the pane has started (settings loaded)
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
  expect(log.inserted[0].data).not.toContain("Max Mustermann"); // no signature in the card – it is in the password mail
  expect(log.customProps.cloudshareJob).toBe("job-1"); // links the draft to the password mail
  expect(JSON.parse(log.session["cloudshare-state"])).toMatchObject({ jobId: "job-1", separate: true });

  // One password per mail: the next link joins this one ("Gleiches Kennwort wie zuvor"); nothing else can be chosen –
  // "Zufällig", "Eigenes", "Ohne Kennwort" and "Kennwort an" are greyed out.
  await expect(page.locator("#mode-same")).toBeChecked();
  await expect(page.locator("#password")).toHaveValue(password);
  await expect(page.locator("#password-hint")).toHaveText("Diese Nachricht hat schon ein Kennwort. Weitere Dateien, Nachrichten und Upload-Links erhalten dasselbe Kennwort.");
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"]).lastStep.password).toBe(password);
  for (const mode of ["#mode-random", "#mode-own", "#mode-none"]) await expect(page.locator(mode)).toBeDisabled();
  // "Kennwort an" the same way: "Gleiche Empfänger wie zuvor" with its hint, the three ways greyed out.
  await expect(page.locator("#channel-same")).toBeChecked();
  await expect(page.locator("#channel-same-row")).toHaveText("Gleiche Empfänger wie zuvor");
  await expect(page.locator("#channel-same-hint")).toHaveText(
    "Diese Nachricht hat schon Empfänger für das Kennwort. Weitere Dateien, Nachrichten und Upload-Links erhalten dieselben Empfänger.");
  await expect(page.locator("#password-channel")).not.toHaveAttribute("title"); // no hover text – the hints say it
  for (const way of ["recipients", "mail", "sms"] as const) {
    if (await channelRadio(page, way).count()) await expect(channelRadio(page, way)).toBeDisabled();
  }
  await expect(page.locator("#password-to-hint")).toBeHidden();
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

test("one password per mail: every further file, message or upload link joins the first step", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Angebot.pdf", 1.5 * MiB), attachment("a2", "Plan.zip", 1.5 * MiB)],
  });
  await expect(page.locator("#submit")).toHaveText("2 Dateien als Link einfügen");
  await expect(page.locator("#mode-same-row")).toBeHidden(); // nothing to share yet
  await expect(page.locator("#password-channel input:checked")).toBeEnabled();
  await page.click("#submit"); // click 1: both files together
  await expect(page.locator("#status")).toContainText("2 Links wurden eingefügt.");

  // Click 2: a confidential message – "Gleiches Kennwort wie zuvor" is the only choice left.
  await expect(page.locator("#mode-same")).toBeChecked();
  await expect(page.locator("#mode-random")).toBeDisabled();
  await expect(page.locator("#mode-own")).toBeDisabled();
  await expect(page.locator("#channel-same")).toBeChecked(); // "Gleiche Empfänger wie zuvor" …
  await expect(channelRadio(page, "recipients")).toBeDisabled(); // … the ways greyed out
  await page.locator("label").filter({ has: page.locator("#kind-message") }).click();
  await page.fill("#message", "Zugangsdaten folgen");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("eingefügt");

  // Click 3: another file – still the same password.
  await page.locator("label").filter({ has: page.locator("#kind-files") }).click();
  await expect(page.locator("#mode-same")).toBeChecked();
  await page.setInputFiles("#file", { name: "Nachtrag.pdf", mimeType: "application/pdf", buffer: Buffer.alloc(2000, 0x41) });
  await page.click("#submit");
  await expect.poll(() => calls.filter((c) => c.path.endsWith("/complete")).length).toBe(3);
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");

  const files = calls.filter((c) => c.path.endsWith("/complete")).map((c) => JSON.parse(c.body!.toString()));
  const message = JSON.parse(calls.find((c) => c.path === "/api/messages")!.body!.toString());
  // One step, one password, one password mail for the whole mail.
  expect(new Set([...files.map((f) => f.step), message.step]).size).toBe(1);
  expect(new Set([...files.map((f) => f.password), message.password]).size).toBe(1);
  expect(files[2].jobId).toBe("job-1");
  await expectAccessible(page);
});

test("Kennwort an eine andere E-Mail-Adresse: stored on the job once it exists", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await expect(page.locator("#password-to")).toBeHidden();
  // Every way has its hint right below it – also "Empfänger der Nachricht".
  await expect(page.locator("#password-to-hint")).toHaveText(
    "Das Kennwort geht an alle Empfänger der Nachricht.");
  await expect(page.locator('#password-channel label:has(input[value="recipients"]) + #password-to-hint')).toHaveCount(1);
  await chooseChannel(page, "mail");
  await expect(page.locator("#password-to-hint")).toContainText("Nur diese Adresse erhält das Kennwort");

  await page.fill("#password-to", "keine adresse");
  await page.press("#password-to", "Tab");
  await expect(page.locator("#status")).toContainText("gültige E-Mail-Adresse");

  await page.fill("#password-to", "kollege@example.com");
  await page.press("#password-to", "Tab");
  await expect(page.locator("#status")).toContainText("nur an kollege@example.com");
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"]).passwordTo).toBe("kollege@example.com");
  await expectAccessible(page);

  // The address travels with the link (per insertion) – no change on the job.
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString())).toMatchObject({ passwordTo: "kollege@example.com" });
  expect(calls.filter((c) => c.method === "PATCH")).toEqual([]);
});

test("Kennwort per SMS: not offered without seven.io on the server", async ({ page }) => {
  await openTaskpane(page);
  await expect(page.locator("#password-channel label.radio:not(#channel-same-row)")).toHaveText(["Empfänger der Nachricht", "Andere E-Mail-Adresse"]);
  await expect(channelRadio(page, "sms")).toHaveCount(0); // no SMS without seven.io on the server
});

test("Kennwort per SMS: number required, normalized to E.164 and stored on the job", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)], config: { smsEnabled: true, smsDefaultCountry: "49" } });
  await chooseChannel(page, "sms");
  await expect(page.locator("#password-sms")).toBeVisible();
  // Both ways of writing a number are accepted – both are shown as examples.
  await expect(page.locator("#password-sms")).toHaveAttribute("placeholder", "0170 1234567 oder +49 170 1234567");
  await expect(page.locator("#password-to-hint")).toContainText("per SMS an diese Nummer");

  // No number yet: the button stays greyed out, as in Outlook.
  await expect(page.locator("#submit")).toBeDisabled();

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
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString())).toMatchObject({ passwordSms: "+491701234567" });
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"])).toMatchObject({ passwordSms: "+491701234567" });
});

test("Kennwort per SMS: numbers of other countries and service numbers are refused right away", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)],
    config: { smsEnabled: true, smsCountryCodes: ["49", "43", "41"], smsCountries: "Deutschland, Österreich und Schweiz" },
  });
  await chooseChannel(page, "sms");
  // The hint stays short – the countries are named when a number from elsewhere is entered.
  await expect(page.locator("#password-to-hint")).toHaveText("Das Kennwort geht nach dem Senden per SMS an diese Nummer, die Empfänger erhalten keine Kennwort-Nachricht.");

  await page.fill("#password-sms", "+44 7700 900123");
  await page.press("#password-sms", "Tab");
  await expect(page.locator("#status")).toContainText("nur an Rufnummern aus Deutschland, Österreich und Schweiz");
  await expect(page.locator("#password-sms")).toHaveAttribute("aria-invalid", "true");

  await expect(page.locator("#submit")).toBeDisabled(); // no link with a number the password could never reach
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
  await expect(page.locator("#submit")).toBeEnabled();
});

test("Kennwort an: suggestions from the own contacts while typing – arrow keys and Enter, or a click", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)],
    config: { smsEnabled: true, smsCountryCodes: ["49"], smsCountries: "Deutschland" },
    contacts: [
      { name: "Jonas Müller-Berg", value: "j.mueller@example.org" },
      { name: "Anna Hamüller", value: "anna@example.test" },
      { name: "Jonas Müller-Berg", value: "+49 171 1234567", type: "mobile" },
      { name: "Jonas Müller-Berg", value: "+44 7700 900123", type: "business" }, // no SMS to that country
    ],
  });
  await chooseChannel(page, "mail");
  const field = page.locator("#password-to");
  await field.fill("mül");
  const list = page.locator("#password-to-suggestions");
  await expect(list).toBeVisible();
  await expect(list).toHaveAttribute("aria-label", "Vorschläge aus Ihren Kontakten");
  await expect(field).toHaveAttribute("aria-expanded", "true");
  await expect(list.locator("[role=option] .suggestion-name")).toHaveText(["Jonas Müller-Berg", "Anna Hamüller"]);
  await expect(list.locator("[role=option] .suggestion-value")).toHaveText(["j.mueller@example.org", "anna@example.test"]);
  await expectAccessible(page);
  await field.press("ArrowDown");
  await expect(list.locator("[role=option]").nth(0)).toHaveAttribute("aria-selected", "true");
  await field.press("Enter");
  await expect(field).toHaveValue("j.mueller@example.org");
  await expect(list).toBeHidden();
  await expect(field).toHaveAttribute("aria-expanded", "false");
  await expect(page.locator("#status")).toContainText("nur an j.mueller@example.org");
  await expect(page.locator("#submit")).toBeEnabled();

  // Escape closes only the list.
  await field.fill("anna");
  await expect(list).toBeVisible();
  await field.press("Escape");
  await expect(list).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();

  // SMS: only numbers that may get one, a click fills in the number as E.164.
  await chooseChannel(page, "sms");
  const sms = page.locator("#password-sms");
  await sms.fill("jonas");
  const numbers = page.locator("#password-sms-suggestions [role=option]");
  await expect(numbers).toHaveCount(1);
  await expect(numbers.locator(".suggestion-value")).toHaveText("+49 171 1234567 · Mobil");
  await numbers.nth(0).click();
  await expect(sms).toHaveValue("+491711234567");
  await expect(page.locator("#status")).toContainText("per SMS an +491711234567");
  expect(calls.filter((c) => c.path === "/api/contacts").length).toBeGreaterThan(1);
});

test("Kennwort an: the address book next to the field – all contacts A to Z, searchable, a click takes one", async ({ page }) => {
  const calls = await openTaskpane(page, {
    attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)],
    config: { smsEnabled: true, smsCountryCodes: ["49"], smsCountries: "Deutschland" },
    contacts: [
      { name: "Anna Beispiel", value: "anna@example.com" },
      { name: "Bernd Muster", value: "bernd@example.org" },
      { name: "Bernd Muster", value: "+49 171 1234567", type: "mobile" },
      { name: "Bernd Muster", value: "+44 7700 900123", type: "business" }, // no SMS to that country
    ],
  });
  await chooseChannel(page, "mail");
  await expect(page.locator("#password-to-book")).toBeVisible();
  await expect(page.locator("#password-to-book")).toHaveAttribute("aria-label", "Aus dem Adressbuch wählen");
  await page.click("#password-to-book");
  await expect(page.locator("#panel-contacts")).toBeVisible();
  await expect(page.locator("#panel-send")).toBeHidden();
  await expect(page.locator("#contacts-heading")).toHaveText("Adressbuch");
  await expect(page.locator("#contacts-search")).toBeFocused();
  await expect(page.locator(".contacts-letter")).toHaveText(["A", "B"]);
  await expect(page.locator(".contact-entry")).toHaveCount(2);
  expect(calls.find((c) => c.path === "/api/contacts")).toBeTruthy();
  await expectAccessible(page);
  await shot(page, "address-book");

  // The search filters; a click takes the address, saved like a typed one, and leads back.
  await page.fill("#contacts-search", "bernd");
  await expect(page.locator(".contact-entry")).toHaveCount(1);
  await page.fill("#contacts-search", "niemand");
  await expect(page.locator("#contacts-empty")).toHaveText("Kein Kontakt passt zur Suche.");
  await page.fill("#contacts-search", "bern");
  await page.locator(".contact-entry").first().click();
  await expect(page.locator("#panel-contacts")).toBeHidden();
  await expect(page.locator("#password-to")).toHaveValue("bernd@example.org");
  await expect(page.locator("#status")).toContainText("nur an bernd@example.org");

  // SMS: only the numbers that may get one.
  await chooseChannel(page, "sms");
  await page.click("#password-sms-book");
  await expect(page.locator(".contact-entry")).toHaveCount(1);
  await expect(page.locator(".contact-value")).toHaveText("+49 171 1234567 · Mobil");
  await page.locator(".contact-entry").click();
  await expect(page.locator("#password-sms")).toHaveValue("+491711234567");

  // Escape leads back without taking anything.
  await page.click("#password-sms-book");
  await page.keyboard.press("Escape");
  await expect(page.locator("#panel-contacts")).toBeHidden();
  await expect(page.locator("#password-sms")).toHaveValue("+491711234567");
});

test("Kennwort an: without Contacts.Read the address book says so and its buttons disappear", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await chooseChannel(page, "mail");
  await page.click("#password-to-book");
  await expect(page.locator("#status")).toContainText("Ihre Outlook-Kontakte sind für CloudShare nicht freigegeben");
  await expect(page.locator("#panel-contacts")).toBeHidden();
  await expect(page.locator("#password-to-book")).toBeHidden();
  await expect(page.locator("#password-to")).toBeFocused();
});

test("Kennwort an without the permission Contacts.Read: a plain field, asked only once", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await chooseChannel(page, "mail");
  const field = page.locator("#password-to");
  await field.pressSequentially("mül", { delay: 50 });
  await expect.poll(() => calls.filter((c) => c.path === "/api/contacts").length).toBe(1);
  await field.pressSequentially("ler", { delay: 300 });
  await expect(page.locator("#password-to-suggestions")).toBeHidden();
  expect(calls.filter((c) => c.path === "/api/contacts")).toHaveLength(1);
});

test("Kennwort an eine andere E-Mail-Adresse chosen without an address: no link until it is there", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await chooseChannel(page, "mail");

  // No address yet: greyed out for every kind, nothing is shared.
  await expect(page.locator("#submit")).toBeDisabled();
  await chooseKind(page, "request");
  await expect(page.locator("#submit")).toBeDisabled();
  expect(calls.some((c) => c.path === "/api/uploads" || c.path === "/api/requests")).toBe(false); // nothing shared

  // Typed and clicked right away (no "change" yet): the address still travels with the link.
  await chooseKind(page, "files");
  await page.fill("#password-to", "kollege@example.com");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString())).toMatchObject({ passwordTo: "kollege@example.com" });
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
  await expect(page.locator("#channel-same-row")).toBeHidden(); // one password per mail – free again …
  await expect(channelRadio(page, "recipients")).toBeChecked(); // … with the way it had
  await expect(channelRadio(page, "mail")).toBeEnabled();
  const log = await officeLog(page);
  expect(log.customProps.cloudshareJob).toBeUndefined(); // the send handler has nothing to arm
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({});
});

test("Feedback senden: not offered without FEEDBACK_MAIL on the server", async ({ page }) => {
  await openTaskpane(page);
  await page.click("#account-toggle");
  await expect(page.locator("#account-menu")).toBeVisible();
  await expect(page.locator("#feedback-open")).toBeHidden();
});

test("Feedback senden: a problem report with the last error message of the pane", async ({ page }) => {
  const calls = await openTaskpane(page, { config: { feedback: true } });
  await chooseChannel(page, "mail");
  await page.fill("#password-to", "keine-adresse");
  await page.press("#password-to", "Tab");
  await expect(page.locator("#status")).toContainText("gültige E-Mail-Adresse"); // an error the report will carry

  await page.click("#account-toggle");
  await expect(page.locator("#account-menu button:visible")).toContainText(["Einstellungen", "Feedback senden", "Hilfe"]);
  await page.click("#feedback-open");
  await expect(page.locator("#panel-feedback")).toBeVisible();
  await expect(page.locator("#panel-send")).toBeHidden();
  await expect(page.locator("#feedback-text")).toBeFocused();
  await expect(page.locator("#feedback-kind-idea")).toBeChecked();
  await expect(page.locator("#feedback-diagnostics")).toBeChecked();
  await expectAccessible(page);

  // Nothing written yet: not sent.
  await page.click("#feedback-send");
  await expect(page.locator("#status")).toHaveText("Bitte schreiben Sie zuerst Ihre Nachricht.");
  expect(calls.some((c) => c.path === "/api/feedback")).toBe(false);

  // The details are shown before they go.
  await page.locator("label.segment", { has: page.locator("#feedback-kind-bug") }).click();
  await page.fill("#feedback-text", "Der Link landet mitten in der Karte.");
  await page.click("#feedback-details-toggle");
  await expect(page.locator("#feedback-details")).toBeVisible();
  await expect(page.locator("#feedback-details")).toContainText("Letzte Fehlermeldungen");
  await expect(page.locator("#feedback-details")).toContainText("gültige E-Mail-Adresse");
  await expect(page.locator("#feedback-details-toggle")).toHaveText("Angaben ausblenden");
  await page.click("#feedback-send");
  await expect(page.locator("#status")).toHaveText("Danke! Ihre Nachricht wurde gesendet.");
  const sent = JSON.parse(calls.find((c) => c.path === "/api/feedback")!.body!.toString());
  expect(sent).toMatchObject({ kind: "bug", text: "Der Link landet mitten in der Karte." });
  expect(sent.diagnostics.errors).toContain("gültige E-Mail-Adresse");
  expect(Object.keys(sent.diagnostics)).toEqual(expect.arrayContaining(["outlook", "addinLanguage", "browser", "time"]));
  // Back on "Senden", the text emptied for the next time.
  await expect(page.locator("#panel-feedback")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
  await expect(page.locator("#feedback-text")).toHaveValue("");

  // Without the details nothing technical goes along.
  await page.click("#account-toggle");
  await page.click("#feedback-open");
  await page.fill("#feedback-text", "Bitte mehr Farben.");
  await page.locator("label.checkbox", { has: page.locator("#feedback-diagnostics") }).click();
  await expect(page.locator("#feedback-details-toggle")).toBeHidden();
  await page.click("#feedback-send");
  await expect(page.locator("#status")).toHaveText("Danke! Ihre Nachricht wurde gesendet.");
  const second = JSON.parse(calls.filter((c) => c.path === "/api/feedback")[1].body!.toString());
  expect(second).toEqual({ kind: "idea", text: "Bitte mehr Farben." });
});

test("a further link goes right after the last card – wherever the cursor is", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Bericht.pdf", 1.5 * MiB)] });
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect((await officeLog(page)).inserted[0].at).toBe("cursor"); // the first card: at the cursor
  await chooseKind(page, "request");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link zum Hochladen wurde eingefügt.");
  const log = await officeLog(page);
  // The body was written back once with the new card between the first card and the empty line after it.
  expect(log.inserted).toHaveLength(1);
  expect(log.inserted[0].at).toBe("body");
  const body = log.inserted[0].data;
  const first = body.indexOf("https://cloud.example.com/s/u1");
  const second = body.indexOf("https://cloud.example.com/s/request");
  expect(first).toBeGreaterThan(-1);
  expect(second).toBeGreaterThan(first);
  expect(body.lastIndexOf("</table>", second)).toBeGreaterThan(first); // after the first card's frame, not inside it
  expect(body.endsWith("<p style=\"margin:0;\"><br></p>")).toBe(true);
  expect(JSON.parse(log.session["cloudshare-state"]).links).toEqual(["u1", "request"]);
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
  await expect(status).toContainText("Verlauf");
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
  await expect(rows.nth(2).locator("input")).toHaveAttribute("aria-label", "Drei.pdf");
  await rows.nth(0).locator(".remove-file").click(); // remove "Eins.pdf"
  await expect(rows).toHaveCount(2);
  await expect(rows.nth(0)).toContainText("Zwei.pdf");
  await expect(rows.nth(0).locator("input")).toBeChecked(); // the × did not toggle anything
  await expect(rows.nth(1)).toContainText("Drei.pdf");
  await expect(rows.nth(1).locator("input")).not.toBeChecked(); // stays unticked after re-rendering
  await expect(rows.nth(0).locator(".remove-file")).toHaveAttribute("aria-label", "„Zwei.pdf“ entfernen");
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
  // The field sits right below the chosen option: under "Zufälliges Kennwort" first, then under "Eigenes Kennwort".
  await expect(page.locator("label:has(#mode-random) + #password-group")).toHaveCount(1);
  await expect(page.locator("#password-group")).toBeHidden(); // a random password is not shown
  await page.check("#mode-own");
  // Confirmed at the top like a new way of "Kennwort an" – never with the password.
  await expect(page.locator("#status")).toHaveText("Sie vergeben das Kennwort selbst.");
  await page.check("#mode-random");
  await expect(page.locator("#status")).toHaveText("Das Kennwort wird zufällig erzeugt.");
  await expect(page.locator("#status")).not.toContainText(await page.locator("#password").inputValue());
  await page.check("#mode-own");
  await expect(page.locator("label:has(#mode-own) + #password-group + #password-rules + #password-hint")).toHaveCount(1);
  await expect(page.locator("#password-group")).toBeVisible(); // an own one needs its field
  const password = page.locator("#password");
  await expect(password).toHaveValue("");
  // Visible while typing (no eye) and never type="password" – Edge would offer to save it as a login.
  await expect(password).toHaveAttribute("type", "text");
  await expect(password).not.toHaveClass(/\bmasked\b/);
  await expect(page.locator("#password-group button")).toHaveCount(0);
  await expect(page.locator("#password-rules")).toBeVisible();

  await password.fill("abc");
  await expect(page.locator('#password-rules li[data-rule="length"]')).toHaveClass(/fail/);
  await expect(page.locator("#submit")).toBeDisabled(); // greyed out until the rules are met
  expect(calls.some((c) => c.path === "/api/uploads")).toBe(false);

  await password.fill(STRONG);
  for (const rule of ["length", "classes", "variety"]) {
    await expect(page.locator(`#password-rules li[data-rule="${rule}"]`)).toHaveClass(/ok/);
  }
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
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({ noPassword: true, links: ["u1"], proofMail: false });

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
  await expect(page.locator("#submit")).toBeDisabled(); // the attachment is a link now – nothing selected
  await page.setInputFiles("#file", { name: "Nachtrag.pdf", mimeType: "application/pdf", buffer: Buffer.alloc(2000, 0x41) });
  await expect(page.locator("#submit")).toBeEnabled(); // further files without password are fine
});

test("Nachricht / Anfordern: 'Ohne Kennwort' is not offered – a choice made for files becomes a random password", async ({ page }) => {
  await openTaskpane(page);
  await page.click("label:has(#mode-none)");
  await expect(page.locator("#password-group")).toBeHidden();
  await chooseKind(page, "request");
  await expect(page.locator("#mode-none-row")).toBeHidden();
  // Greyed out instead of gone, the reason right below it (no "i" to click).
  const unavailable = page.locator("#mode-none-unavailable");
  await expect(unavailable).toBeVisible();
  await expect(unavailable.locator("input")).toBeDisabled();
  await expect(unavailable.locator("button")).toHaveCount(0);
  await expect(page.locator("#mode-none-why")).toBeVisible();
  await expect(page.locator("#mode-none-why")).toHaveText(
    "Nur für Dateien: Eine vertrauliche Nachricht und ein Upload-Link sind immer mit einem Kennwort geschützt.");
  await expectAccessible(page);
  await expect(page.locator("#mode-random")).toBeChecked();
  await expect(page.locator("#password-group")).toBeHidden(); // a random password is not shown – it goes out by itself
  await expect(page.locator("#password")).not.toHaveValue("");
  await expect(page.locator("#submit")).toHaveText("Link zum Hochladen einfügen");
  await chooseKind(page, "files");
  await expect(page.locator("#mode-none-row")).toBeVisible();
  await expect(unavailable).toBeHidden();
  await expect(page.locator("#submit")).toHaveText("Als Link einfügen"); // no attachments, nothing ticked
});

test("Nextcloud enforces link passwords: no 'Ohne Kennwort' option", async ({ page }) => {
  await openTaskpane(page, { passwordEnforced: true });
  await expect(page.locator("#account")).toContainText("max.mustermann"); // connection status has arrived
  await expect(page.locator("label:has(#mode-none)")).toBeHidden();
  await expect(page.locator("#mode-random")).toBeChecked();
  await expect(page.locator("#password-group")).toBeHidden(); // the random password is not shown
});

test("Dateien von Empfängern erhalten inserts an upload link with the same password", async ({ page }) => {
  const calls = await openTaskpane(page);
  await expect(page.locator("#password")).not.toHaveValue(""); // the pane has started (settings loaded)
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

test("Zustellnachweis: not offered without DELIVERY_PROOF on the server", async ({ page }) => {
  // Without DELIVERY_PROOF on the server the box does not exist for the user.
  await openTaskpane(page);
  await page.click("#validity-toggle");
  await expect(page.locator("#proof-field")).toBeHidden();
  await chooseKind(page, "message");
  await expect(page.locator("#proof-field")).toBeHidden();
});

test("Zustellnachweis: only for files; ticked, it goes with the upload and the card names the tracking ID", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Plan.pdf", 2 * MiB)], config: { deliveryProof: true } });
  await expect(page.locator("#password")).not.toHaveValue("");
  // Below the password, in sight (no heading "Empfänger": "Kennwort an" belongs to the password) – not there at all for a
  // message or an upload link.
  await expect(page.locator("#panel-send h2", { hasText: /^Empfänger$/ })).toHaveCount(0);
  expect(await page.locator(".password-field + #password-to-field + #proof-field").count()).toBe(1);
  await expect(page.locator("#password-to-field h2.section-title")).toHaveText("Kennwort an"); // a heading like "Kennwort"
  for (const kind of ["message", "request"] as const) {
    await chooseKind(page, kind);
    await expect(page.locator("#proof-field")).toBeHidden();
  }
  await expectAccessible(page);
  await chooseKind(page, "files");
  await expect(page.locator("#proof-field")).toBeVisible();
  await expect(page.locator("#proof")).not.toBeChecked();
  await expect(page.locator("#proof")).toHaveAttribute("role", "switch");
  await page.locator("#proof-field .switch-text").click();
  await expect(page.locator("#proof")).toBeChecked();
  await expectAccessible(page);

  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  const body = JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString());
  expect(body).toMatchObject({ proof: true, client: { host: expect.any(String) } });
  const card = (await officeLog(page)).inserted[0].data as string;
  expect(card).toContain("https://cloudshare.example.com/r/s/u"); // the gateway, not the Nextcloud link
  expect(card).toContain("Vorgang 20261117-1146-7Q2F");
  // The password comes after the confirmation – no "separate message" with the link mail.
  expect(card).toContain("Das Kennwort erhalten Sie, nachdem Sie Ihre E-Mail-Adresse bestätigt haben.");
  expect(card).not.toContain("separaten Nachricht");

  // From now on this mail takes only files with a proof: no confidential message, no upload link, the switch fixed on.
  await expect(page.locator("#proof")).toBeChecked();
  await expect(page.locator("#proof")).toBeDisabled();
  // No box at the top, no "i": the reason follows the text under the switch.
  await expect(page.locator("#proof-field button")).toHaveCount(0);
  await expect(page.locator("#proof-hint")).toHaveText(
    "Versand und Abrufe werden protokolliert; Empfänger bestätigen ihre Adresse mit einem Code. " +
    "Diese Nachricht hat einen Zustellnachweis – weitere Dateien gehen ebenfalls mit Nachweis. " +
    "Vertrauliche Nachricht und Upload-Link bitte in einer eigenen Nachricht.");
  await expect(page.locator('input[name="kind"][value="message"]')).toBeDisabled();
  await expect(page.locator('input[name="kind"][value="request"]')).toBeDisabled();
  await expect(page.locator('label.segment:has(input[value="message"])')).not.toHaveAttribute("title"); // no hover text
  expect(JSON.parse((await officeLog(page)).session["cloudshare-state"])).toMatchObject({ proofMail: true });
});

test("Zustellnachweis: after a link without it, the mail takes no proof (never mixed in one mail)", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Plan.pdf", 2 * MiB)], config: { deliveryProof: true } });
  await expect(page.locator("#password")).not.toHaveValue("");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect(page.locator("#proof")).not.toBeChecked();
  await expect(page.locator("#proof")).toBeDisabled();
  await expect(page.locator("#proof-hint")).toHaveText(
    "Versand und Abrufe werden protokolliert; Empfänger bestätigen ihre Adresse mit einem Code. " +
    "Diese Nachricht enthält schon Links ohne Zustellnachweis. Für einen Zustellnachweis bitte eine eigene Nachricht verwenden.");
  // A message or an upload link is still fine here.
  await expect(page.locator('input[name="kind"][value="message"]')).toBeEnabled();
  await expectAccessible(page);
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

  await expect(page.locator("#password")).not.toHaveValue(""); // the pane has started (settings loaded)
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

  // Empty text: the button is greyed out until there is a message.
  await expect(page.locator("#submit")).toBeDisabled();
  await page.fill("#message", "Noch eine");
  await expect(page.locator("#submit")).toBeEnabled();
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
  await chooseChannel(page, "sms");
  await expect(page.locator("#max-access-field")).toBeHidden();
  await expect(page.locator("#password-days-field")).toBeHidden();
  await expect(page.locator("#password-to-hint")).toContainText("per SMS an diese Nummer");
  await page.fill("#password-sms", "0170 1234567");
  await page.locator("#password-sms").blur();
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  // The password link only holds the password until sending – as long as possible.
  expect(JSON.parse(calls.find((c) => c.path.endsWith("/complete"))!.body!.toString()).passwordDays).toBe(30);

  // The mail has its password now: "Kennwort an" is fixed for the next links (one password, one way).
  await expect(page.locator("#channel-same")).toBeChecked(); // "Gleiche Empfänger wie zuvor" …
  await expect(channelRadio(page, "recipients")).toBeDisabled(); // … the ways greyed out
});

test("Ohne Kennwort with download confirmation: the draft is tagged so the confirmation knows the sent mail", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Flyer.pdf", 1.5 * MiB)], noticeJob: true });
  await page.click("label:has(#mode-none)");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("ohne Kennwort abrufbar");
  const log = await officeLog(page);
  expect(log.customProps.cloudshareJob).toBe("notice-1");
  expect(JSON.parse(log.session["cloudshare-state"])).toEqual({ noPassword: true, jobId: "notice-1", links: ["u1"], proofMail: false });
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

// --- History ---------------------------------------------------------------------------------------------------

const SHARES = [
  { id: "11", kind: "file", name: "Präsentation.pptx", url: "https://cloud.example.com/s/a", sharedAt: Date.parse("2026-11-10"),
    expiration: "2099-11-24", downloads: { count: 2, last: "2026-11-11T09:30:00Z" }, received: null, protected: true,
    password: { channel: "sms", to: "+49 170 1234567", status: "delivered", at: Date.parse("2026-11-10T10:00:00Z"), deliveredAt: Date.parse("2026-11-10T10:00:09Z") } },
  { id: "12", kind: "file", name: "Angebot.pdf", url: "https://cloud.example.com/s/b", sharedAt: Date.parse("2099-11-12"),
    expiration: "2099-11-26", downloads: { count: 0, last: null }, received: null, protected: false },
  { id: "13", kind: "request", name: "Upload-Link", url: "https://cloud.example.com/s/c", sharedAt: Date.parse("2026-11-01"),
    expiration: "2099-11-15", downloads: null, received: 3 },
  { id: "14", kind: "file", name: "Alt.zip", url: "https://cloud.example.com/s/d", sharedAt: Date.parse("2020-01-01"),
    expiration: "2020-01-15", downloads: { count: 0, last: null }, received: null },
];


test("Verlauf: a link with a proof of delivery shows its tracking ID under the file, the seal after it", async ({ page }) => {
  await openTaskpane(page, { shares: [{ ...SHARES[0], proofId: "20261110-0912-7Q2F", proofState: "fetched" }, SHARES[1]] });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  // As in "Zustellnachweise": the number alone (bold) right under the file, the seal for the PDF after it.
  await expect(rows.nth(0).locator(".share-item + .proof-case")).toHaveText("20261110-0912-7Q2F");
  await expect(rows.nth(0).locator(".proof-case .proof-seal")).toHaveClass(/state-fetched/);
  await expect(rows.nth(0).locator(".share-item .proof-seal")).toHaveCount(0);
  // Without a proof: no such line.
  await expect(rows.nth(1).locator(".proof-case")).toHaveCount(0);
  await expectAccessible(page);
});
test("Verlauf: shows the state of each link and revokes after confirmation", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  // All valid links, no expired one ("Alt.zip"), no "last 30 days".
  await expect(rows).toHaveCount(3);
  await expect(page.locator("#shares")).not.toContainText("Alt.zip");
  // The head of the page: "←", "Aktualisieren" and the magnifier (the list also loads when opened and again on focus).
  await expect(page.locator("#panel-shares .help-header button")).toHaveCount(3);
  expect(calls.filter((c) => c.path === "/api/shares").length).toBeGreaterThan(0);
  // One quiet bottom bar for every tab: only the account, as a menu (like the menus of Outlook and OneDrive).
  const toggle = page.locator("#account-toggle");
  const menu = page.locator("#account-menu");
  const storage = page.locator("#account-storage");
  await expect(toggle).toHaveText("max.mustermann");
  await expect(menu).toBeHidden();
  await expect(storage).toHaveText("37,5 GB von 50 GB belegt");
  await expect(toggle).toHaveAttribute("title", "Angemeldet als max.mustermann");
  // What the menu holds: for screen readers only (the arrow shows it on screen).
  await expect(toggle).toHaveAccessibleDescription("Menü: Nextcloud öffnen, Einstellungen, Hilfe, Verbindung trennen");
  const toggleBox = (await toggle.boundingBox())!;
  await expect(storage).toBeHidden(); // the storage is in the menu now, nothing else in the bar
  await expect(page.locator(".account-bar")).toHaveCount(1);
  // The menu opens above the bar: each entry with an icon; Escape closes it, focus back on the account.
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  await expect(menu).toBeVisible();
  await expect(page.locator("#account-menu button:visible")).toHaveText(["Nextcloud öffnen", "Verlauf", "Einstellungen", "Hilfe", "Verbindung trennen"]);
  await expect(page.locator("#account-menu button:visible svg")).toHaveCount(5);
  await expect(page.locator("#nextcloud-open")).toBeFocused(); // the first entry
  await expect(page.locator("#nextcloud-open")).toHaveAttribute("title", "https://cloud.example.com");
  // At the top, set apart: the whole name, a bar filled with the used share (37,5 of 50 GB = 75 %) and the text.
  await expect(page.locator("#account-menu-name")).toHaveText("max.mustermann");
  await expect(storage).toBeVisible();
  await expect(page.locator("#storage-meter")).toBeVisible();
  await expect(page.locator("#storage-meter-fill")).toHaveAttribute("style", "width: 75%;");
  await expect(page.locator("#storage-meter")).toHaveClass(/\bwarn\b/); // orange from 75 %, red from 90 %
  await expect(page.locator("#storage-meter")).not.toHaveClass(/\bfull\b/);
  expect((await page.locator(".account-menu-head").boundingBox())!.y).toBeLessThan((await page.locator("#nextcloud-open").boundingBox())!.y);
  expect((await menu.boundingBox())!.y + (await menu.boundingBox())!.height).toBeLessThanOrEqual(toggleBox!.y);
  await expectAccessible(page);
  await page.keyboard.press("Escape");
  await expect(menu).toBeHidden();
  await expect(toggle).toBeFocused();
  // A click elsewhere closes it too.
  await toggle.click();
  await page.locator("#shares").click({ position: { x: 5, y: 5 } });
  await expect(menu).toBeHidden();
  await page.click("#shares-close");
  await expect(page.locator("#panel-shares")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
  await openHistory(page);
  await expect(rows.nth(0).locator(".state-text")).toHaveText("2× heruntergeladen · 11.11.2026, 10:30 Uhr");
  await expect(rows.nth(1).locator(".state-text")).toHaveText("Noch nicht heruntergeladen");
  await expect(rows.nth(2).locator(".state-text")).toHaveText("3 Dateien empfangen");
  // The same lines as an entry of "Zustellnachweise": the day (and to whom, once sent), "Kennwort an …", "Gültig bis …"
  // with its calendar and the state alone on the last line.
  await expect(rows.nth(0).locator(".share-sent")).toHaveText("10.11.2026");
  await expect(rows.nth(0).locator(".share-password")).toHaveText("Kennwort an +49 170 1234567 – zugestellt");
  await expect(rows.nth(0).locator(".share-password")).toHaveAttribute("title", "Kennwort per SMS an +49 170 1234567 – zugestellt am 10.11.2026, 11:00 Uhr");
  await expect(rows.nth(0).locator(".share-valid")).toHaveText("Gültig bis 24.11.2099");
  await expect(rows.nth(0).locator(".share-valid")).not.toHaveClass(/soon/);
  await expect(rows.nth(0).locator(".share-state")).not.toContainText("24.11.2099"); // no end date in the state line
  expect(await rows.nth(0).locator(".share-valid + .share-state").count()).toBe(1); // right above the state
  await expect(rows.nth(1).locator(".share-sent")).toHaveText("12.11.2099"); // always dd.mm.yyyy
  await expect(rows.nth(1).locator(".share-password")).toHaveText("Ohne Kennwort");
  await expect(rows.nth(2).locator(".share-password")).toHaveText("Mit Kennwort");
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

test("an error message runs down too; one with \"Erneut versuchen\" stays", async ({ page }) => {
  await openTaskpane(page, {
    attachments: [attachment("a1", "Bericht.pdf", 1.5 * MiB)],
    uploadError: { status: 507, error: "Nicht genügend Speicherplatz in Ihrer Nextcloud.", code: "NEXTCLOUD_QUOTA_EXCEEDED" },
    config: { successMessageSeconds: 1 },
  });
  await page.click("#submit");
  const status = page.locator("#status");
  await expect(status).toHaveClass(/error/);
  await expect(status.locator(".status-timer")).toHaveCount(1); // the red ring around the "!"
  await page.mouse.move(5, 700);
  await expect(status).toBeHidden({ timeout: 3000 });
});

test("a failed start keeps its message with \"Erneut versuchen\", whatever the timer", async ({ page }) => {
  await openTaskpane(page, { tokenError: { code: 13007, name: "Error", message: "x" }, config: { successMessageSeconds: 1 } });
  const status = page.locator("#status");
  await expect(status.locator("button.status-action")).toBeVisible();
  await expect(status.locator(".status-timer")).toHaveCount(0);
  await page.waitForTimeout(1500);
  await expect(status).toBeVisible();
});

test("a success message runs down and goes; the pointer on it stops the ring", async ({ page }) => {
  await openTaskpane(page, { shares: SHARES, config: { successMessageSeconds: 1 } });
  await openHistory(page);
  const revoke = page.locator("#shares .share").nth(1).locator(".revoke");
  await revoke.click();
  await revoke.click();
  const status = page.locator("#status");
  await expect(status).toHaveClass(/success/);
  await expect(status.locator(".status-timer")).toHaveCount(1); // the ring around the tick
  // Held with the pointer: it stays past its second.
  await status.hover();
  await page.waitForTimeout(1500);
  await expect(status).toBeVisible();
  // Let go: it runs down and goes.
  await page.mouse.move(5, 700);
  await expect(status).toBeHidden({ timeout: 3000 });
});

test("Verlauf: a failed SMS says whether it was sent at all, and how to send the password again", async ({ page }) => {
  const sms = { channel: "sms", to: "+49 170 1234567", status: "failed", at: Date.parse("2026-11-10T10:00:00Z"), deliveredAt: null };
  const share = (id: string, name: string, refused: boolean) => ({ ...SHARES[0], id, name, url: `https://cloud.example.com/s/${id}`,
    password: { ...sms, refused }, passwordAvailable: true });
  await openTaskpane(page, { shares: [share("21", "Refused.pdf", true), share("22", "Lost.pdf", false)] });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  // seven.io refused it right away (key, credit, limit).
  await expect(rows.nth(0).locator(".share-password")).toHaveText("Kennwort an +49 170 1234567 – nicht versendet – bitte erneut senden");
  // seven.io accepted it and reported "not delivered" later.
  await expect(rows.nth(1).locator(".share-password")).toHaveText("Kennwort an +49 170 1234567 – nicht zugestellt – bitte erneut senden");
  await expect(rows.nth(1).locator(".share-password")).toHaveClass(/failed/);
});

test("Verlauf: links of a mail not sent (e.g. a discarded draft) say so and when they are revoked", async ({ page }) => {
  const revokeAt = Date.parse("2099-11-22T10:00:00Z");
  await openTaskpane(page, {
    shares: [
      { ...SHARES[0], id: "31", password: null, unsent: true, revokeAt, downloads: { count: 0, last: null } },
      { ...SHARES[0], id: "32", password: null, unsent: true, revokeAt: null, downloads: { count: 0, last: null } },
      // Downloaded although CloudShare has not seen the mail sent (e.g. opened by the sender): only the password is missing.
      { ...SHARES[0], id: "33", password: null, unsent: true, revokeAt: null },
      SHARES[1],
    ],
  });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  await expect(rows.nth(0).locator(".share-unsent")).toHaveText(
    `Nicht gesendet – wird die Nachricht bis 22.11.2099 nicht gesendet, zieht CloudShare den Link zurück.`,
  );
  await expect(rows.nth(1).locator(".share-unsent")).toHaveText("Nicht gesendet – das Kennwort wurde noch nicht verschickt.");
  await expect(rows.nth(2).locator(".share-unsent")).toHaveText("Kennwort noch nicht verschickt.");
  await expect(rows.nth(2).locator(".state-text")).toContainText("2× heruntergeladen");
  await expect(rows.nth(3).locator(".share-unsent")).toBeHidden();
  // Not sent: the bold line says it – no extra "Kennwort noch nicht gesendet" line above it.
  await expect(rows.nth(0).locator(".share-password")).toBeHidden();
  await expect(rows.nth(3).locator(".share-password")).toBeVisible();
  await expectAccessible(page);
});

test("Verlauf: Kennwort shows the password and sends it again the same ways as the first time", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: [{ ...SHARES[0], passwordAvailable: true }, SHARES[1]], config: { smsEnabled: true } });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  await expect(rows.nth(1).locator(".show-password")).toBeHidden(); // no password to show (link without password)
  const toggle = rows.nth(0).locator(".show-password");
  await expect(toggle).toHaveAttribute("aria-label", "Kennwort von „Präsentation.pptx“ anzeigen und erneut senden");
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  const panel = rows.nth(0).locator(".password-panel");
  await expect(panel.locator("input.mono")).toHaveValue("KateGeteFi45");
  // Hidden at first; the eye shows it (and hides it again); copying works without showing it.
  await expect(panel.locator("input.mono")).toHaveClass(/\bmasked\b/);
  await expect(panel.locator(".reveal-password")).toHaveAttribute("aria-label", "Kennwort anzeigen");
  await panel.locator(".reveal-password").click();
  await expect(panel.locator("input.mono")).not.toHaveClass(/\bmasked\b/);
  await expect(panel.locator(".reveal-password")).toHaveAttribute("aria-pressed", "true");
  await panel.locator(".reveal-password").click();
  await expect(panel.locator("input.mono")).toHaveClass(/\bmasked\b/);
  // Nothing ticked at first – the password goes again only where the sender ticks it; "Senden" waits for a box.
  await expect(panel.locator("input[type=checkbox]")).toHaveCount(5);
  await expect(panel.locator("input[type=checkbox]:checked")).toHaveCount(0);
  await expect(panel.locator(".send-password")).toBeDisabled();
  // SMS ticked: its number field right under its box, the number of the first time filled in.
  const smsField = panel.locator("input[type=tel]");
  await expect(smsField).toBeHidden();
  await panel.locator('label.checkbox:has(input[value="sms"])').click();
  await expect(smsField).toHaveValue("+491701234567");
  await expect(panel.locator('label.checkbox:has(input[value="sms"]) + .channel-group')).toHaveCount(1);
  await expect(panel.locator(".send-password")).toBeEnabled();
  await expectAccessible(page);
  await shot(page, "password-panel");

  // As boxes, so several at once: the recipients of the message one by one (To, Cc, Bcc below each, nothing to
  // confirm), another address, SMS.
  await expect(panel.locator("label.checkbox .checkbox-text")).toHaveText([
    /^erika@example\.com\s*An$/, /^chef@example\.com\s*Cc$/, /^archiv@example\.com\s*Bcc$/, /^Andere E-Mail-Adresse/, /^Rufnummer \(SMS\)/]);
  // Another address: its field only while ticked (one address, like "Kennwort an" when sending).
  const mailField = panel.locator("input[type=email]");
  await expect(mailField).toBeHidden();
  await panel.locator('label.checkbox:has(input[value="mail"])').click();
  await expect(mailField).toBeVisible();
  await expect(mailField).toHaveValue("");
  await expect(mailField).toHaveAttribute("placeholder", "name@beispiel.de");
  // Nothing typed: told, nothing sent.
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toContainText("Bitte die E-Mail-Adresse angeben");
  await mailField.fill("kollege@example.com");
  // Two of the three recipients.
  await panel.locator('label.checkbox:has(input[value="erika@example.com"])').click();
  await panel.locator('label.checkbox:has(input[value="archiv@example.com"])').click();

  // All at once: the ticked recipients in one request (one mail, each in its field), then the address, then SMS.
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toContainText("Das Kennwort wurde an kollege@example.com geschickt.");
  const sends = calls.filter((c) => c.path === "/api/shares/11/password/send").map((c) => JSON.parse(c.body!.toString()));
  expect(sends).toEqual([
    { channel: "recipients", to: ["erika@example.com", "archiv@example.com"] },
    { channel: "mail", to: "kollege@example.com" }, { channel: "sms", to: "+491701234567" }]);
  // The list reloads quietly after sending (new state of the password in the row); the panel stays as it is – not
  // closed, not loaded again –, only unticked, "Senden" greyed out, the fields under the boxes folded away.
  await expect.poll(() => calls.filter((c) => c.path === "/api/shares").length).toBe(2);
  await expect(panel).toBeVisible();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  expect(calls.filter((c) => c.path === "/api/shares/11/password").length).toBe(1);
  await expect(panel.locator("input[type=checkbox]")).toHaveCount(5);
  await expect(mailField).toBeHidden();
  await expect(panel.locator("input[type=checkbox]:checked")).toHaveCount(0);
  await expect(panel.locator(".send-password")).toBeDisabled();
});

test("Verlauf: Erneut senden an – contact suggestions and the address book fill in the one address and lead back", async ({ page }) => {
  await openTaskpane(page, {
    shares: [{ ...SHARES[0], passwordAvailable: true }],
    config: { smsEnabled: true },
    contacts: [
      { name: "Anna Beispiel", value: "anna@example.com" },
      { name: "Bernd Muster", value: "bernd@example.org" },
    ],
  });
  await openHistory(page);
  const row = page.locator("#shares .share").nth(0);
  await row.locator(".show-password").click();
  const panel = row.locator(".password-panel");
  await panel.locator('label.checkbox:has(input[value="mail"])').click();
  const field = panel.locator("input[type=email]");
  const book = panel.locator(".channel-group:has(input[type=email]) .book-open");
  await expect(book).toBeVisible();
  await expect(book).toHaveAttribute("aria-label", "Aus dem Adressbuch wählen");

  // Typing: suggestions from the contacts, a click takes the address (one address, no list).
  await field.fill("ann");
  const options = panel.locator(".suggestions [role=option]");
  await expect(options).toHaveCount(1);
  await options.nth(0).click();
  await expect(field).toHaveValue("anna@example.com");

  // The address book: opened from "Verlauf", a click replaces the address and leads back to "Verlauf", the panel still open.
  await book.click();
  await expect(page.locator("#panel-contacts")).toBeVisible();
  await expect(page.locator("#panel-shares")).toBeHidden();
  await page.locator(".contact-entry", { hasText: "Bernd Muster" }).click();
  await expect(page.locator("#panel-contacts")).toBeHidden();
  await expect(page.locator("#panel-shares")).toBeVisible();
  await expect(field).toHaveValue("bernd@example.org");
  await expect(field).toBeFocused();

  // Escape on the address book also leads back to "Verlauf", not to "Senden".
  await book.click();
  await page.keyboard.press("Escape");
  await expect(page.locator("#panel-shares")).toBeVisible();
  await expect(field).toHaveValue("bernd@example.org");

  // Unticked, the address field goes – the address book button with it.
  await panel.locator('label.checkbox:has(input[value="mail"])').click();
  await expect(book).toBeHidden();
  await expectAccessible(page);
});

test("Verlauf: a password held back for the proof of delivery is not sent again before a confirmation", async ({ page }) => {
  const calls = await openTaskpane(page, {
    shares: [{ ...SHARES[0], passwordAvailable: true, proofId: "20261117-1146-7Q2F" }],
    passwordInfo: { deferred: true, verified: [], password: undefined, sms: "" },
  });
  await openHistory(page);
  const row = page.locator("#shares .share").nth(0);
  await expect(row.locator(".show-password")).toHaveAttribute("aria-label", "Kennwort von „Präsentation.pptx“ erneut senden");
  await row.locator(".show-password").click();
  const panel = row.locator(".password-panel");
  await expect(panel.locator("input")).toHaveCount(0); // no password, no field, no way to pick
  await expect(panel.locator(".reveal-password, .copy-password")).toHaveCount(0);
  await expect(panel.locator(".label")).toHaveText("Kennwort"); // nothing to send yet: no "Erneut senden an"
  await expect(panel.locator(".hint")).toHaveText(
    "Wird beim Zustellnachweis nicht angezeigt. Jeder Empfänger erhält es automatisch, sobald er seine Adresse " +
    "bestätigt – erneut senden können Sie es erst danach.");
  await expect(panel.locator(".send-password")).toHaveCount(0);
  await expectAccessible(page);
  expect(calls.some((c) => c.path.endsWith("/password/send"))).toBe(false);
});

test("Verlauf: with a proof of delivery the password goes again only to the confirmed recipients ticked", async ({ page }) => {
  const calls = await openTaskpane(page, {
    shares: [{ ...SHARES[0], passwordAvailable: true, proofId: "20261117-1146-7Q2F" }],
    passwordInfo: { deferred: true, verified: ["erika@example.com", "bernd@example.com"], password: undefined, sms: "" },
  });
  await openHistory(page);
  const row = page.locator("#shares .share").nth(0);
  await row.locator(".show-password").click();
  const panel = row.locator(".password-panel");
  await expect(panel.locator("input.mono")).toHaveCount(0);
  await expect(panel.locator(".hint").first()).toHaveText(
    "Das Kennwort wird nicht angezeigt – nur erneut gesendet, an Empfänger, die ihre Adresse " +
    "bestätigt haben. Jeder erhält eine eigene Nachricht.");
  const boxes = panel.locator("label.checkbox");
  // "An", "Cc" or "Bcc" on the right as without a proof (bernd is not in the mocked distribution: nothing there).
  await expect(boxes).toHaveText([/^erika@example\.com\s*An$/, "bernd@example.com"]);
  await expect(panel.locator("input[type=checkbox]:checked")).toHaveCount(0); // nobody gets it unasked
  const send = panel.locator(".send-password");
  await expect(send).toBeDisabled();
  await boxes.nth(1).click();
  await expect(send).toBeEnabled();
  await expectAccessible(page);
  await send.click();
  await expect(page.locator("#status")).toContainText("Das Kennwort wurde an");
  const sends = calls.filter((c) => c.path.endsWith("/password/send")).map((c) => JSON.parse(c.body!.toString()));
  expect(sends).toEqual([{ channel: "recipients", to: ["bernd@example.com"] }]);
  // Still open after sending, as it was – only unticked.
  await expect.poll(() => calls.filter((c) => c.path === "/api/shares").length).toBe(2);
  await expect(panel).toBeVisible();
  await expect(boxes).toHaveCount(2);
  await expect(panel.locator("input[type=checkbox]:checked")).toHaveCount(0);
});

test("Verlauf: many recipients fold into one box – ticked all at once or one by one from the list with a filter", async ({ page }) => {
  const to = ["erika@example.com", "max@example.com", "anna@example.com", "bernd@example.com", "maria@example.com"];
  const calls = await openTaskpane(page, {
    shares: [{ ...SHARES[0], passwordAvailable: true }],
    passwordInfo: { recipients: { to, cc: ["chef@example.com"], bcc: [] } },
  });
  await openHistory(page);
  const panel = page.locator("#shares .share").nth(0).locator(".password-panel");
  await page.locator("#shares .share").nth(0).locator(".show-password").click();
  const all = panel.locator("label.checkbox:has(.recipient-all)");
  await expect(all.locator(".checkbox-text")).toHaveText(/^Empfänger der Nachricht \(6\)\s*Keiner ausgewählt$/);
  const list = panel.locator(".recipient-list");
  await expect(list).toBeHidden(); // one line until needed
  await expect(panel.locator(".send-password")).toBeDisabled();

  // The arrow folds out the list, one line per address with "An"/"Cc" on the right; the filter narrows it.
  await panel.locator(".recipient-toggle").click();
  await expect(panel.locator(".recipient-toggle")).toHaveAttribute("aria-expanded", "true");
  await expect(list.locator(".recipient-filter")).toBeFocused();
  await expect(list.locator("label.recipient-row")).toHaveCount(6);
  await expect(list.locator('label.recipient-row:has(input[value="chef@example.com"]) .recipient-field')).toHaveText("Cc");
  await list.locator(".recipient-filter").fill("ma");
  await expect(list.locator("label.recipient-row:visible")).toHaveCount(2); // max, maria
  await list.locator('label.recipient-row:has(input[value="maria@example.com"])').click();
  await expect(all.locator(".checkbox-hint")).toHaveText("1 von 6 ausgewählt");
  await expect(panel.locator(".recipient-all")).toHaveJSProperty("indeterminate", true);
  await expect(panel.locator(".send-password")).toBeEnabled();
  await list.locator(".recipient-filter").fill("xyz");
  await expect(list.locator(".recipient-none")).toBeVisible();
  await list.locator(".recipient-filter").fill("");
  await expectAccessible(page);

  // The box of the group: everyone.
  await all.click();
  await expect(all.locator(".checkbox-hint")).toHaveText("Alle ausgewählt");
  await expect(list.locator("input:checked")).toHaveCount(6);
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toContainText("Das Kennwort wurde an");
  const sends = calls.filter((c) => c.path.endsWith("/password/send")).map((c) => JSON.parse(c.body!.toString()));
  expect(sends).toEqual([{ channel: "recipients", to: [...to, "chef@example.com"] }]);
  // Afterwards the list stays folded out as it was, everyone unticked again.
  await expect(all.locator(".checkbox-hint")).toHaveText("Keiner ausgewählt");
  await expect(list).toBeVisible();
  await expect(panel.locator(".send-password")).toBeDisabled();
});

test("Verlauf: an open password panel stays open when the list is built anew", async ({ page }) => {
  // The first click into the pane focuses it, and focus reloads the list – that must not close the panel just opened.
  await openTaskpane(page, {
    shares: [{ ...SHARES[0], passwordAvailable: true, proofId: "20261117-1146-7Q2F" }],
    passwordInfo: { deferred: true, verified: ["erika@example.com"], password: undefined, sms: "" },
  });
  await openHistory(page);
  const row = page.locator("#shares .share").nth(0);
  await row.locator(".show-password").click();
  await row.locator(".password-panel label.checkbox").first().click();
  await page.click("#shares-search-toggle");
  await page.fill("#shares-search", "Präsentation"); // builds the list anew
  await expect(page.locator("#shares .password-panel")).toBeVisible();
  await expect(page.locator("#shares .password-panel input[type=checkbox]").first()).toBeChecked();
});

test("Verlauf: with a proof of delivery the password may go by SMS to a corrected number", async ({ page }) => {
  const calls = await openTaskpane(page, {
    config: { smsEnabled: true },
    shares: [{ ...SHARES[0], passwordAvailable: true, proofId: "20261117-1146-7Q2F" }],
    passwordInfo: { deferred: true, verified: ["erika@example.com"], password: undefined, sms: "+491701234567" },
  });
  await openHistory(page);
  const row = page.locator("#shares .share").nth(0);
  await row.locator(".show-password").click();
  const panel = row.locator(".password-panel");
  const number = panel.locator("input[type=tel]");
  await expect(number).toBeHidden(); // only once "Rufnummer (SMS)" is ticked
  await panel.locator("label.checkbox", { hasText: "Rufnummer (SMS)" }).click();
  await expect(number).toBeVisible();
  await expect(number).toHaveValue("+491701234567");
  await number.fill("0171 2345678");
  await expectAccessible(page);
  await panel.locator(".send-password").click();
  await expect(page.locator("#status")).toContainText("Das Kennwort wurde an");
  const sends = calls.filter((c) => c.path.endsWith("/password/send")).map((c) => JSON.parse(c.body!.toString()));
  expect(sends).toEqual([{ channel: "sms", to: "0171 2345678" }]);
});

test("Verlauf: Als Link einfügen creates a new link with this mail's password", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await expect(page.locator("#password")).not.toHaveValue(""); // the pane has started (settings loaded)
  const password = await page.locator("#password").inputValue();
  await openHistory(page);
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

test("Verlauf: the subject is the title and opens the mail; confidential messages with password and revoke only", async ({ page }) => {
  const message = { id: "msg-s1", kind: "message", name: "Vertrauliche Nachricht", url: null, accessId: "acc9", sharedAt: Date.parse("2026-11-12"),
    expiration: "2099-11-17", downloads: null, received: null, protected: true, passwordAvailable: true, subject: "Zugangsdaten", mailId: "AAMk-mail",
    sentAt: Date.parse("2026-11-12T09:00:00Z"), sentTo: { to: ["Erika@example.com"], cc: [] },
    password: { channel: "mail", to: "erika@example.com", status: null, at: Date.parse("2026-11-12T08:00:00Z"), deliveredAt: null } };
  const calls = await openTaskpane(page, {
    shares: [
      { ...SHARES[0], subject: "Unterlagen Neubau", mailId: "AAMk-files", sentAt: Date.parse("2026-11-16T09:04:00Z"),
        sentTo: { to: ["erika@example.com", "max@example.com"], cc: ["chef@example.com"] } },
      { ...SHARES[1], subject: "", mailId: null }, // no subject, mail in a shared mailbox: not openable
      message,
      SHARES[2], // not sent yet: no subject known – titled with its name as before
    ],
  });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  await expect(rows.nth(0).locator(".share-title")).toHaveText("Unterlagen Neubau");
  // Hovering the subject: when and to whom it went; screen readers hear that it opens the mail.
  await expect(rows.nth(0).locator(".share-title")).toHaveAttribute(
    "title",
    "Gesendet am 16.11.2026, 10:04 Uhr an erika@example.com, max@example.com\nCc: chef@example.com",
  );
  await expect(rows.nth(0).locator(".share-title")).toHaveAttribute("aria-label", "Unterlagen Neubau, Nachricht öffnen");
  // Like "Zustellnachweise": the day it went out and to whom (To and Cc) right in the entry.
  await expect(rows.nth(0).locator(".share-sent")).toHaveText("16.11.2026 · Nachricht an erika@example.com, max@example.com, chef@example.com");
  await expect(rows.nth(0).locator(".share-password")).toHaveText("Kennwort an +49 170 1234567 – zugestellt"); // SMS: its own line
  // Below the subject what went out, in plain text: the file name (the document symbol on the left shows the kind).
  await expect(rows.nth(0).locator(".share-item")).toHaveText("Präsentation.pptx");
  await expect(rows.nth(0).locator(".badge")).toHaveCount(0);
  await expect(rows.nth(0).locator(".share-item .share-name")).toHaveAttribute("title", "Präsentation.pptx");
  await expect(rows.nth(1).locator(".share-title")).toHaveText("(Kein Betreff)");
  await expect(rows.nth(1).locator("button.share-title")).toHaveCount(0);
  // An entry from before subjects were kept: the same layout, "(Betreff unbekannt)" as its title.
  await expect(rows.nth(3).locator(".share-title")).toHaveText("(Betreff unbekannt)");
  await expect(rows.nth(3).locator(".share-item")).toHaveText("Dateianforderung");

  // A click on the subject opens the sent mail in Outlook (the server gives its current EWS ID).
  await rows.nth(0).locator("button.share-title").click();
  await expect.poll(async () => (await officeLog(page)).openedMails).toEqual(["ews-11"]); // translated by the server

  // The message: envelope, "Vertrauliche Nachricht" below the subject, the password – but no copy, no insert, no
  // download state (CloudShare keeps no key, Vaultwarden reports no opening).
  const msg = rows.nth(2);
  await expect(msg.locator(".share-title")).toHaveText("Zugangsdaten");
  await expect(msg.locator(".share-item")).toHaveText("Vertrauliche Nachricht"); // only the label, in colour
  await expect(msg.locator(".show-password")).toBeVisible();
  await expect(msg.locator(".copy")).toBeHidden();
  await expect(msg.locator(".insert")).toBeHidden();
  await expect(msg.locator(".share-state")).toBeHidden(); // no download state for a message
  await expect(msg.locator(".share-valid")).toHaveText("Gültig bis 17.11.2099");
  // Message and password to the same address (case does not matter): one line, the password line hidden.
  await expect(msg.locator(".share-sent")).toHaveText("12.11.2026 · Nachricht und Kennwort an Erika@example.com");
  await expect(msg.locator(".share-password")).toBeHidden();
  await expectAccessible(page);

  // Revoked: deleted in Vaultwarden.
  await msg.locator(".revoke").click();
  await msg.locator(".revoke").click();
  await expect(page.locator("#status")).toHaveText("„Vertrauliche Nachricht“ ist nicht mehr abrufbar. Die Nachricht wurde gelöscht.");
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/msg-s1"]);
  await expect(rows).toHaveCount(3);
});

test("Verlauf: a password sent again names when it last went out – also to the same addresses", async ({ page }) => {
  const sentTo = { to: ["erika@example.com"], cc: [] };
  await openTaskpane(page, {
    shares: [
      { ...SHARES[0], sentAt: Date.parse("2026-11-12T09:00:00Z"), sentTo,
        password: { channel: "mail", to: "erika@example.com", status: null, at: Date.parse("2026-11-14T13:32:00Z"), deliveredAt: null, resent: true } },
      { ...SHARES[1], protected: true, sentAt: Date.parse("2026-11-12T09:00:00Z"), sentTo,
        password: { channel: "sms", to: "+49 170 1234567", status: "delivered", at: Date.parse("2026-11-15T08:05:00Z"),
          deliveredAt: Date.parse("2026-11-15T08:05:09Z"), resent: true } },
    ],
  });
  await openHistory(page);
  const mail = page.locator("#shares .share", { hasText: "Präsentation.pptx" });
  const sms = page.locator("#shares .share", { hasText: "Angebot.pdf" });
  // Not "Nachricht und Kennwort an …": the message line names the message alone, the password its own line with the time.
  await expect(mail.locator(".share-sent")).toHaveText("12.11.2026 · Nachricht an erika@example.com");
  await expect(mail.locator(".share-password")).toHaveText("Kennwort zuletzt an erika@example.com · 14.11.2026, 14:32 Uhr");
  await expect(sms.locator(".share-password")).toHaveText("Kennwort zuletzt an +49 170 1234567 · 15.11.2026, 09:05 Uhr – zugestellt");
});

test("Verlauf and Zustellnachweise: many recipients – the first two and \"und N weitere\", all of them in the tooltip", async ({ page }) => {
  const many = ["erika@example.com", "max@example.com", "anna@example.com", "bernd@example.com", "maria@example.com"];
  await openTaskpane(page, {
    config: { deliveryProof: true },
    shares: [{ ...SHARES[0], sentAt: Date.parse("2026-11-12T09:00:00Z"), sentTo: { to: many.slice(0, 4), cc: many.slice(4) },
      password: { channel: "mail", to: many.join(", "), status: null, at: Date.parse("2026-11-12T09:05:00Z"), deliveredAt: null } }],
    proofs: [{ id: "20261112-1000-AB12", createdAt: Date.parse("2026-11-12T09:00:00Z"), subject: "Rundschreiben", sentAt: Date.parse("2026-11-12T09:00:00Z"),
      recipients: many, files: ["Plan.pdf"], downloads: 0, confirmed: 0, done: false, state: "sent", passwordTo: ["kollege@example.com"],
      endsOn: "2099-12-01", keptUntil: Date.parse("2029-11-12T12:00:00Z"), mailAvailable: false, kind: "file", fetched: 0 }],
  });
  await openHistory(page);
  const sent = page.locator("#shares .share").nth(0).locator(".share-sent");
  await expect(sent).toHaveText("12.11.2026 · Nachricht und Kennwort an erika@example.com, max@example.com und 3 weitere");
  expect(await sent.getAttribute("title")).toContain(many.join(", "));
  // The search still finds a recipient that is not written out.
  await page.click("#shares-search-toggle");
  await page.fill("#shares-search", "maria@");
  await expect(page.locator("#shares .share")).toHaveCount(1);

  await page.click("#shares-close");
  await page.click("#account-toggle");
  await page.click("#proofs-open");
  const proofSent = page.locator("#proofs-list .proof-sent").first();
  await expect(proofSent).toHaveText("12.11.2026 · Nachricht an erika@example.com, max@example.com und 3 weitere");
  await expect(proofSent).toHaveAttribute("title", many.join(", "));
});

test("reading a mail: many recipients shortened in its links too – \"und N weitere\", all of them in the tooltip", async ({ page }) => {
  const many = ["erika@example.com", "max@example.com", "anna@example.com", "bernd@example.com", "maria@example.com"];
  await openTaskpane(page, {
    readBody: SENT_MAIL,
    shares: [{ ...SHARES[0], sentAt: Date.parse("2026-11-12T09:00:00Z"), sentTo: { to: many, cc: [] },
      password: { channel: "mail", to: many.slice(1).join(", "), status: null, at: Date.parse("2026-11-12T09:05:00Z"), deliveredAt: null } }],
  });
  const row = page.locator("#mail-shares .share").first();
  await expect(row.locator(".share-sent")).toHaveText("12.11.2026 · Nachricht an erika@example.com, max@example.com und 3 weitere");
  expect(await row.locator(".share-sent").getAttribute("title")).toContain(many.join(", "));
  await expect(row.locator(".share-password")).toHaveText("Kennwort an max@example.com, anna@example.com und 2 weitere");
});

test("reading a mail: a confidential message of that mail shows under 'Nachricht'", async ({ page }) => {
  const message = { id: "msg-r1", kind: "message", name: "Vertrauliche Nachricht", url: null, accessId: "accR1", sharedAt: Date.parse("2026-11-12"),
    expiration: "2099-11-17", downloads: null, received: null, protected: true, subject: "Zugangsdaten", mailId: "AAMk-r" };
  await openTaskpane(page, {
    shares: [SHARES[1], message],
    readBody: '<p>Hallo,</p><a href="https://vault.example.com/#/send/accR1/geheimerSchluessel">Nachricht öffnen</a>',
  });
  await expect(page.locator("#mail-shares .share")).toHaveCount(1);
  await expect(page.locator("#mail-shares .share .share-item")).toHaveText("Vertrauliche Nachricht");
});

test("reading a mail: the links of that mail, \"Verlauf\" in the menu, nothing to insert", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES, readBody: SENT_MAIL });
  // No "Senden" and no tabs when reading: the links of this mail; "Verlauf" in the account menu as when writing.
  await expect(page.locator("#panel-mail")).toBeVisible();
  await expect(page.locator("#mail-heading")).toHaveText("Links in dieser Nachricht");
  await expect(page.locator("#account-toggle")).toBeVisible(); // "Verlauf", the guide and the settings are in its menu
  await expect(page.locator("#panel-send")).toBeHidden();
  await expect(page.locator("#panel-shares")).toBeHidden();
  const inMail = page.locator("#mail-shares .share");
  await expect(inMail).toHaveCount(2);
  await expect(inMail.nth(0).locator(".share-name")).toHaveText("Präsentation.pptx");
  await expect(inMail.nth(0).locator(".share-password")).toHaveText("Kennwort an +49 170 1234567 – zugestellt");
  await expect(inMail.nth(0).locator(".share-password")).toHaveAttribute("title", /zugestellt am 10.11.2026, 11:00 Uhr/);
  await expect(inMail.nth(1).locator(".share-item")).toHaveText("Dateianforderung");
  await expect(page.locator("#mail-none")).toBeHidden();
  // Reading a mail too: "Aktualisieren" right of the title (it also reloads by itself), the account at the bottom.
  await expect(page.locator("#panel-mail .list-head button")).toHaveCount(1);
  await expect(page.locator("#mail-refresh")).toHaveAttribute("aria-label", "Aktualisieren");
  await expect(page.locator("#account-toggle")).toHaveText("max.mustermann");
  await expect(page.locator(".share .insert").first()).toBeHidden(); // a mail being read cannot take a link
  await expect(page.locator("#panel-mail")).not.toContainText("Weitere Freigaben");
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "read-view");

  // "Verlauf": every valid link, those of this mail included (Alt.zip has expired).
  await openHistory(page);
  await expect(page.locator("#panel-mail")).toBeHidden();
  await expect(page.locator("#shares .share .share-item")).toHaveText(["Präsentation.pptx", "Angebot.pdf", "Dateianforderung"]);
  await expectAccessible(page);

  // Revoked in one list, gone from both.
  await page.click("#shares-close");
  await expect(page.locator("#panel-mail")).toBeVisible();
  const revoke = page.locator("#mail-shares .share").nth(1).locator(".revoke");
  await revoke.click();
  await revoke.click();
  await expect(page.locator("#mail-shares .share")).toHaveCount(1);
  await expect(page.locator("#shares .share")).toHaveCount(2);
  expect(calls.filter((c) => c.method === "DELETE").map((c) => c.path)).toEqual(["/api/shares/13"]);
  await expect(page.locator("#status")).toBeVisible(); // "… ist nicht mehr abrufbar …"

  // Pinned pane, the user selects another mail without CloudShare links: back to "Nachricht", the message about
  // the previous mail is gone.
  await openHistory(page);
  await page.evaluate(() => {
    (window as unknown as { __officeConfig: { readBody: string } }).__officeConfig.readBody = "<p>Keine Links</p>";
    (window as unknown as { __officeHandlers: Record<string, () => void> }).__officeHandlers.olkItemSelectedChanged();
  });
  await expect(page.locator("#panel-mail")).toBeVisible();
  await expect(page.locator("#panel-shares")).toBeHidden();
  await expect(page.locator("#status")).toBeHidden();
  await expect(page.locator("#mail-none")).toHaveText("Diese Nachricht enthält keine CloudShare-Links von Ihnen.");
  await expect(page.locator("#mail-shares .share")).toHaveCount(0);
});

test("reading a mail before connecting Nextcloud: connect view, then the shares", async ({ page }) => {
  await openTaskpane(page, { connected: false, shares: SHARES, readBody: SENT_MAIL });
  await expect(page.locator("#connect-view")).toBeVisible();
  await page.click("#connect");
  await expect(page.locator("#mail-shares .share")).toHaveCount(2, { timeout: 10_000 });
  await expect(page.locator("#panel-send")).toBeHidden();
});

test("Verlauf: a page of the account menu with a search over subject, file and recipient", async ({ page }) => {
  await openTaskpane(page, { shares: SHARES });
  await openHistory(page);
  const rows = page.locator("#shares .share");
  await expect(rows).toHaveCount(3);
  // The search sits behind the magnifier in the head, like in "Hilfe".
  await expect(page.locator("#shares-search")).toBeHidden();
  await expect(page.locator("#shares-search-toggle")).toHaveAttribute("aria-label", "Verlauf durchsuchen");
  await page.click("#shares-search-toggle");
  await expect(page.locator("#shares-search-toggle")).toHaveAttribute("aria-expanded", "true");
  await expect(page.locator("#shares-search")).toBeFocused();
  await page.fill("#shares-search", "angebot");
  await expect(rows).toHaveCount(1);
  await expect(rows.nth(0).locator(".share-item")).toContainText("Angebot.pdf");
  await page.fill("#shares-search", "nichts dergleichen");
  await expect(rows).toHaveCount(0);
  await expect(page.locator("#no-shares")).toHaveText("Kein Eintrag passt zur Suche.");
  await expectAccessible(page);
  // Escape closes the search only (not the page) and brings the whole list back.
  await page.locator("#shares-search").press("Escape");
  await expect(page.locator("#shares-search")).toBeHidden();
  await expect(page.locator("#panel-shares")).toBeVisible();
  await expect(rows).toHaveCount(3);
});

test("Verlauf: upload links are not inserted again without password", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES });
  await page.check("#mode-none");
  await openHistory(page);
  await page.locator("#shares .share").nth(2).locator(".insert").click();
  await expect(page.locator("#status")).toContainText("Links zum Hochladen brauchen immer ein Kennwort");
  expect(calls.some((c) => c.path.endsWith("/link"))).toBe(false);
});

// --- Help -----------------------------------------------------------------------------------------------------------

test("Hilfe: \"Hilfe\" in the account menu opens the guide in place of the tab, readable at 320 px", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await openTaskpane(page);
  await expectNoHorizontalScroll(page); // tabs and the bottom bar (account and storage) fit
  const toggle = page.locator("#help-toggle");
  await openHelp(page);
  await expect(page.locator("#account-menu")).toBeHidden();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  await expect(page.locator("#panel-help")).toBeVisible();
  await expect(page.locator("#help-close")).toBeFocused(); // screen readers start in the guide
  await expect(page.locator("#help-heading")).toHaveText("Hilfe");
  await expect(page.locator("#panel-send")).toBeHidden();
  const help = page.locator("#panel-help");
  await expect(help).not.toContainText("help."); // every text exists in the language file
  // "Zustellnachweis" only where the server offers it (DELIVERY_PROOF) – hidden here.
  const sections = help.locator("details:not([hidden])");
  await expect(sections).toHaveCount(8);
  // The most common task is open, the rest closed until needed.
  await expect(sections.nth(1).locator("summary")).toHaveText("Große Dateien senden");
  await expect(sections.nth(1)).toHaveAttribute("open", "");
  await expect(sections.nth(7)).not.toHaveAttribute("open", "");
  await sections.nth(7).locator("summary").click();
  await expect(sections.nth(7)).toHaveAttribute("open", "");
  await expect(sections.nth(7)).toContainText("Ich habe den Link an die falsche Person geschickt.");
  // "Einstellungen" says where to find them: in the account menu at the bottom left.
  const settings = sections.filter({ has: page.locator("summary", { hasText: "Einstellungen" }) });
  await settings.locator("summary").click();
  await expect(settings).toContainText("auf Ihr Konto und dann auf „Einstellungen“");
  await expect(settings).toContainText("Zurücksetzen");
  // "Verlauf" explains every icon with the icon itself – "Aktualisieren" too.
  const shares = sections.filter({ has: page.locator("summary", { hasText: "Verlauf" }) });
  await shares.locator("summary").click();
  for (const icon of ["refresh", "copy", "password", "insert", "revoke"]) {
    await expect(shares.locator(`[data-icon="${icon}"] svg`)).toHaveCount(1);
  }
  await expect(shares).toContainText("jederzeit mit „Aktualisieren“ oben rechts");
  await expect(shares).toContainText("„Verbindung trennen“ – Trennen braucht zur Sicherheit einen zweiten Klick");
  await expect(help).toContainText("Personenbezogene oder vertrauliche Daten versenden Sie bitte immer mit Kennwort.");
  await expectAccessible(page);
  await expectNoHorizontalScroll(page);
  await shot(page, "help");

  // "×" closes it: back where the user was, focus on the account.
  await page.click("#help-close");
  await expect(page.locator("#panel-help")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
  await expect(toggle).toHaveAttribute("aria-expanded", "false");
  await expect(page.locator("#account-toggle")).toBeFocused();

  // Opened while "Verlauf" is open: the guide takes its place, back leads to "Senden".
  await openHistory(page);
  await openHelp(page);
  await expect(page.locator("#panel-shares")).toBeHidden();
  await page.click("#help-close");
  await expect(page.locator("#panel-help")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
});

test("Hilfe: the magnifier searches the guide – only sections with a hit, opened and marked; Escape brings it back", async ({ page }) => {
  await openTaskpane(page, { config: { smsEnabled: true } });
  await openHelp(page);
  const sections = page.locator("#panel-help details.help-section:not([hidden])");
  const all = await sections.count();
  const toggle = page.locator("#help-search-toggle");
  await expect(toggle).toHaveAttribute("aria-label", "In der Hilfe suchen");
  await expect(page.locator("#help-search")).toBeHidden();
  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-expanded", "true");
  await expect(page.locator("#help-search")).toBeFocused();

  // Every word must be in a section; the others go, the hits are marked (case and umlauts do not matter).
  await page.locator("#help-search").fill("SMS zugestellt");
  const shown = page.locator("#panel-help details.help-section:not([hidden]):not(.search-miss)");
  await expect.poll(() => shown.count()).toBeGreaterThan(0);
  expect(await shown.count()).toBeLessThan(all);
  for (const section of await shown.all()) await expect(section).toHaveAttribute("open", "");
  await expect(page.locator("#panel-help mark.help-hit").first()).toBeVisible();
  await expect(page.locator("#panel-help .help-intro")).toBeHidden();
  await expectAccessible(page);

  // Nothing found: said so.
  await page.locator("#help-search").fill("xyzzy");
  await expect(shown).toHaveCount(0);
  await expect(page.locator("#help-search-none")).toHaveText("Nichts gefunden – versuchen Sie ein anderes Wort.");

  // Escape closes the search only – the guide is back as it was, no marks left, the page stays open.
  await page.locator("#help-search").press("Escape");
  await expect(page.locator("#help-search")).toBeHidden();
  await expect(page.locator("#panel-help")).toBeVisible();
  await expect(shown).toHaveCount(all);
  await expect(page.locator("#panel-help mark.help-hit")).toHaveCount(0);
  await expect(page.locator("#panel-help .help-intro")).toBeVisible();
});

test("pages open at their top; back to the view it stands where it was left", async ({ page }) => {
  await page.setViewportSize({ width: 376, height: 400 }); // short, so the pane scrolls
  await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pdf", 1.5 * MiB)] });
  await expect(page.locator("#panel-send")).toBeVisible();
  await page.evaluate(() => window.scrollTo(0, 150));
  const scrolled = await page.evaluate(() => window.scrollY);
  expect(scrolled).toBeGreaterThan(0);

  // "Hilfe" from the menu: at its top, even though "Senden" was scrolled down.
  await openHelp(page);
  expect(await page.evaluate(() => window.scrollY)).toBe(0);
  await page.evaluate(() => window.scrollTo(0, 300));
  // Another page from the menu: at its top too.
  await openSettings(page);
  expect(await page.evaluate(() => window.scrollY)).toBe(0);
  // Back to "Senden": where it was left.
  await page.keyboard.press("Escape");
  await expect(page.locator("#panel-send")).toBeVisible();
  expect(await page.evaluate(() => window.scrollY)).toBe(scrolled);
});

test("loading: placeholders while the pane starts and while lists and settings load – replaced once they are there", async ({ page }) => {
  await openTaskpane(page, { shares: SHARES, proofs: [], config: { deliveryProof: true }, slowLoads: 1500 });
  // The start: placeholders in the shape of "Senden" (before any answer), read out as "Wird geladen …".
  const start = page.locator("#start-view");
  await expect(start).toBeVisible();
  await expect(start.locator("[role=status]")).toHaveAttribute("aria-busy", "true");
  await expect(start).toContainText("Wird geladen …");
  await expect(start).toBeHidden({ timeout: 10_000 });
  await expect(page.locator("#panel-send")).toBeVisible();

  // "Verlauf": placeholder rows until the list is there, then the entries – no "Noch keine Freigaben" in between.
  await openHistory(page);
  await expect(page.locator("#shares .skeleton")).toBeVisible();
  await expect(page.locator("#no-shares")).toBeHidden();
  await expect(page.locator("#shares .share")).toHaveCount(3, { timeout: 10_000 });
  await expect(page.locator("#shares .skeleton")).toHaveCount(0);

  // "Zustellnachweise" the same way; "Einstellungen" were loaded at the start already and show at once.
  await page.click("#shares-close");
  await page.click("#account-toggle");
  await page.click("#proofs-open");
  await expect(page.locator("#proofs-list .skeleton")).toBeVisible();
  await expect(page.locator("#proofs-list .skeleton")).toHaveCount(0, { timeout: 10_000 });
  await page.keyboard.press("Escape");
  await openSettings(page);
  await expect(page.locator("#settings-form")).toBeVisible();
  await expect(page.locator("#settings-loading")).toBeHidden();
});
test("Aktualisieren: right before the magnifier in Verlauf and Zustellnachweise – turns while it loads, the list stays", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES, proofs: [], config: { deliveryProof: true } });
  const loads = (path: string) => calls.filter((c) => c.method === "GET" && c.path === path).length;
  await openHistory(page);
  await expect(page.locator("#shares .share")).toHaveCount(3);
  const refresh = page.locator("#shares-refresh");
  await expect(page.locator("#panel-shares .help-header .refresh + #shares-search-toggle")).toHaveCount(1);
  await expect(refresh).toHaveAttribute("aria-label", "Aktualisieren");
  const before = loads("/api/shares");
  await refresh.click();
  await expect.poll(() => loads("/api/shares")).toBe(before + 1);
  await expect(refresh).not.toHaveAttribute("aria-busy", "true");
  await expect(page.locator("#shares .share")).toHaveCount(3);
  await expect(page.locator("#shares .skeleton")).toHaveCount(0); // no placeholders on a reload
  await expectAccessible(page);

  await page.click("#shares-close");
  await page.click("#account-toggle");
  await page.click("#proofs-open");
  await expect(page.locator("#proofs-empty")).toBeVisible();
  const proofsBefore = loads("/api/proofs");
  await expect(page.locator("#panel-proofs .help-header .refresh + #proofs-search-toggle")).toHaveCount(1);
  await page.click("#proofs-refresh");
  await expect.poll(() => loads("/api/proofs")).toBe(proofsBefore + 1);
});

test("Aktualisieren: reading a mail, the links of that mail load anew", async ({ page }) => {
  const calls = await openTaskpane(page, { shares: SHARES, readBody: SENT_MAIL });
  await expect(page.locator("#mail-shares .share").first()).toBeVisible();
  const loads = () => calls.filter((c) => c.method === "GET" && c.path === "/api/shares").length;
  const before = loads();
  await page.click("#mail-refresh");
  await expect.poll(loads).toBe(before + 1);
  await expect(page.locator("#mail-shares .share").first()).toBeVisible();
});

test("loading: reading a mail, its links show placeholders until they are loaded", async ({ page }) => {
  await openTaskpane(page, { shares: SHARES, readBody: SENT_MAIL, slowLoads: 1500 });
  await expect(page.locator("#mail-shares .skeleton")).toBeVisible({ timeout: 10_000 });
  await expect(page.locator("#mail-none")).toBeHidden();
  await expect(page.locator("#mail-shares .share").first()).toBeVisible({ timeout: 10_000 });
  await expect(page.locator("#mail-shares .skeleton")).toHaveCount(0);
});

test("Nextcloud öffnen opens the user's Nextcloud in the browser", async ({ page }) => {
  await openTaskpane(page);
  await page.locator("#account-toggle").click();
  await page.locator("#nextcloud-open").click();
  await expect(page.locator("#account-menu")).toBeHidden();
  expect((await officeLog(page)).openedUrls).toEqual(["https://cloud.example.com"]);
});

test("a long account name at 320 px: shortened with \"…\" in the bar, whole in the tooltip and the menu", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  const name = "philipp.rohlmann.mit.einem.sehr.langen.namen@verwaltung.example.com";
  await openTaskpane(page, { loginName: name });
  const toggle = page.locator("#account-toggle");
  await expect(toggle).toHaveText(name); // the whole name for screen readers …
  // … and on hover.
  await expect(toggle).toHaveAttribute("title", `Angemeldet als ${name}`);
  const shown = page.locator("#account");
  expect(await shown.evaluate((el) => el.scrollWidth > el.clientWidth)).toBe(true); // cut short with "…"
  // In the menu the name is whole (wrapped if need be), above the storage.
  await toggle.click();
  await expect(page.locator("#account-menu-name")).toHaveText(name);
  await expect(page.locator("#account-storage")).toHaveText("37,5 GB von 50 GB belegt");
  await expectNoHorizontalScroll(page);
});

test("a grey note runs down too (Verbindung getrennt)", async ({ page }) => {
  await openTaskpane(page, { config: { successMessageSeconds: 1 } });
  await page.locator("#account-toggle").click();
  await page.locator("#disconnect").click();
  await page.locator("#disconnect").click();
  const status = page.locator("#status");
  await expect(status).toHaveText("Die Verbindung zu Nextcloud wurde getrennt.");
  await expect(status.locator(".status-timer")).toHaveCount(1); // grey ring around the "i"
  await page.mouse.move(5, 700);
  await expect(status).toBeHidden({ timeout: 3000 });
});

test("Verbindung trennen takes a second click, like Zurückziehen", async ({ page }) => {
  const calls = await openTaskpane(page);
  const disconnect = page.locator("#disconnect");
  const deletes = () => calls.filter((c) => c.method === "DELETE" && c.path === "/api/nextcloud/connection").length;
  await page.locator("#account-toggle").click();
  // First click: "Wirklich trennen?" in red, the menu stays open, nothing disconnected yet.
  await disconnect.click();
  await expect(disconnect).toHaveText("Wirklich trennen?");
  await expect(disconnect).toHaveClass(/confirm/);
  await expect(page.locator("#account-menu")).toBeVisible();
  expect(deletes()).toBe(0);
  await expectAccessible(page);
  // Closing the menu disarms it.
  await page.keyboard.press("Escape");
  await page.locator("#account-toggle").click();
  await expect(disconnect).toHaveText("Verbindung trennen");
  // Two clicks: disconnected, the connect view shows.
  await disconnect.click();
  await disconnect.click();
  await expect(page.locator("#status")).toHaveText("Die Verbindung zu Nextcloud wurde getrennt.");
  await expect(page.locator("#connect-view")).toBeVisible();
  expect(deletes()).toBe(1);
});

// --- Look ----------------------------------------------------------------------------------------------------------

test("dark Office theme: dark pane, still accessible", async ({ page }) => {
  await openTaskpane(page, { theme: "#1f1f1f", attachments: [attachment("a1", "Präsentation.pptx", 2.5 * MiB)], shares: SHARES, config: { smsEnabled: true } });
  await expect(page.locator("html")).toHaveAttribute("data-theme", "dark");
  await expect(page.locator("#attachments label.list-item")).toHaveCount(1);
  await expectAccessible(page);
  await fullShot(page, "send-dark");
  await openHistory(page);
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

test("reading a mail at 320 px: the title fits", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  await openTaskpane(page, { shares: SHARES, readBody: SENT_MAIL });
  await expect(page.locator("#mail-heading")).toBeVisible();
  await expectNoHorizontalScroll(page);
  await shot(page, "read-320px");
});

test("Verlauf at 320 px: a long name is shortened in the middle, the icons stay in the title line", async ({ page }) => {
  await page.setViewportSize({ width: 320, height: 800 });
  const long = "cloudshare-icon-varianten-mit-sehr-langem-namen und Leerzeichen-final.png";
  await openTaskpane(page, { shares: [{ ...SHARES[0], name: long, passwordAvailable: true }, SHARES[1]] });
  await openHistory(page);
  const row = page.locator("#shares .share").first();
  const name = row.locator(".share-name");
  await expect(name).toHaveText(long); // the whole name for screen readers …
  await expect(name).toHaveAttribute("title", long); // … and as tooltip
  await expect(name.locator(".name-end")).toHaveText("inal.png"); // the extension stays visible
  const start = name.locator(".name-start");
  expect(await start.evaluate((el) => el.scrollWidth > el.clientWidth)).toBe(true); // the "…" is in the start
  // Icons in the line of the title, not wrapped below it.
  const nameBox = (await row.locator(".share-title").boundingBox())!;
  for (const icon of ["copy", "show-password", "insert", "revoke"]) {
    const box = (await row.locator(`.${icon}`).boundingBox())!;
    expect(Math.abs(box.y + box.height / 2 - (nameBox.y + nameBox.height / 2))).toBeLessThan(6);
    expect(box.x).toBeGreaterThan(nameBox.x + nameBox.width - 1);
  }
  // A short name is not split.
  await expect(page.locator("#shares .share").nth(1).locator(".name-end")).toHaveText("");
  await expectNoHorizontalScroll(page);
  await shot(page, "shares-320px-long-name");
});

// --- README pictures (English) -------------------------------------------------------------------------------------
// The README shows the pane in English: same mocks, texts from locales/en.json (attachments "readme-*").

test.describe("README pictures (English)", () => {
  test.use({ locale: "en-GB" });

  const ATTACHMENTS = [attachment("a1", "Presentation.pptx", 6.5 * MiB), attachment("a2", "Notes.txt", 200 * 1024)];
  // With the delivery proof offered: the pictures show its switch, the seal in "History" and its help.
  const README_CONFIG = { smsEnabled: true, smsDefaultCountry: "49", thresholdMb: 5, sendExpireDays: 5, sendMaxExpireDays: 30, deliveryProof: true };
  // With the mails they went out with (subject as title), and a confidential message.
  const SUBJECTS = ["Documents new building project", "Quote 2026-118", "Documents new building project", ""];
  const SHARES_EN = [
    ...SHARES.map((share, i) => ({ ...share, name: ["Presentation.pptx", "Quote.pdf", "File request", "Old.zip"][i], subject: SUBJECTS[i], mailId: `AAMk-${i}`,
      // The first link went out with a delivery proof: the seal after its name.
      ...(i === 0 ? { proofId: "20261110-0912-7Q2F", proofState: "fetched" } : {}) })),
    { id: "msg-1", kind: "message", name: "Confidential message", url: null, accessId: "acc1", sharedAt: Date.parse("2026-11-12"),
      expiration: "2099-11-17", downloads: null, received: null, protected: true, passwordAvailable: true, subject: "Access data", mailId: "AAMk-9",
      password: { channel: "mail", to: "erika.beispiel@example.com", status: null, at: Date.parse("2026-11-12T08:00:00Z"), deliveredAt: null } },
  ];

  test("Send: files", async ({ page }) => {
    await openTaskpane(page, { lang: "en", attachments: ATTACHMENTS, config: README_CONFIG });
    await expect(page.locator("#submit")).toHaveText("Insert 2 files as links");
    await expect(page.locator("#validity-summary")).toHaveText("Link 14 days · Password link 5 days, opens 3 times");
    await expect(page.locator("#account")).toHaveText("max.mustermann");
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

  test("History", async ({ page }) => {
    const shares = SHARES_EN.map((share, i) => (i === 0 ? { ...share, passwordAvailable: true } : share));
    await openTaskpane(page, { lang: "en", shares, config: README_CONFIG });
    await openHistory(page);
    await expect(page.locator("#shares .share")).toHaveCount(4);
    await expect(page.locator("#shares .share").nth(1).locator(".state-text")).toHaveText("Not downloaded yet");
    // "Password" opened at the first link: the password and "Send again to".
    await page.locator("#shares .share").nth(0).locator(".show-password").click();
    await expect(page.locator("#shares .share").nth(0).locator(".password-panel input.mono")).toHaveValue("KateGeteFi45");
    await expectAccessible(page);
    await readmeShot(page, "shares");
  });

  test("Proofs of delivery", async ({ page }) => {
    const proofs = [
      { id: "20261118-1646-CBMH", createdAt: Date.parse("2026-11-18T15:46:00Z"), subject: "Documents new building project",
        sentAt: Date.parse("2026-11-18T15:47:00Z"), recipients: ["erika.beispiel@example.com"], files: ["Presentation.pptx"], downloads: 1,
        confirmed: 1, done: false, state: "fetched", fetched: 1, endsOn: "2026-12-02", keptUntil: Date.parse("2029-11-18T12:00:00Z"), kind: "file", mailAvailable: true, passwordTo: ["erika@example.com"] },
      { id: "20261104-1012-K8MD", createdAt: Date.parse("2026-11-04T09:12:00Z"), subject: "Contract 2026",
        sentAt: Date.parse("2026-11-04T09:13:00Z"), recipients: ["customer@example.org"], files: ["Contract.pdf", "Annex.pdf"], downloads: 2,
        confirmed: 1, done: true, state: "fetched", fetched: 2, endsOn: null, keptUntil: Date.parse("2029-11-18T12:00:00Z"), kind: "file", mailAvailable: true, passwordTo: ["erika@example.com"] },
      { id: "20260921-1533-Z4QN", createdAt: Date.parse("2026-09-21T13:33:00Z"), subject: "Plans",
        sentAt: Date.parse("2026-09-21T13:34:00Z"), recipients: ["architect@example.net"], files: ["Plans.zip"], downloads: 0,
        confirmed: 0, done: true, state: "sent", fetched: 0, endsOn: null, keptUntil: Date.parse("2029-10-05T12:00:00Z"), kind: "file", mailAvailable: true, passwordTo: ["erika@example.com"] },
    ];
    await openTaskpane(page, { lang: "en", config: README_CONFIG, proofs });
    await page.click("#account-toggle");
    await page.click("#proofs-open");
    await expect(page.locator("#proofs-list .proof-entry")).toHaveCount(3);
    await expectAccessible(page);
    await readmeShot(page, "proofs");
  });

  test("Reading a sent mail", async ({ page }) => {
    await openTaskpane(page, { lang: "en", shares: SHARES_EN, config: README_CONFIG, readBody: SENT_MAIL });
    await expect(page.locator("#mail-shares .share")).toHaveCount(2);
    await expect(page.locator("#mail-heading")).toHaveText("Links in this message");
    await readmeShot(page, "read");
  });

  test("Help", async ({ page }) => {
    await openTaskpane(page, { lang: "en", config: README_CONFIG });
    await openHelp(page);
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

// --- Sensitivity labels (Microsoft Purview) ---------------------------------------------------------------

/** A minimal Office file (ZIP, stored) with a sensitivity label in docProps/custom.xml – example IDs only. */
function labelledDocx(labelId: string): string {
  const custom =
    '<?xml version="1.0" encoding="UTF-8"?><Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    `<property name="MSIP_Label_${labelId}_Enabled"><vt:lpwstr>true</vt:lpwstr></property>` +
    `<property name="MSIP_Label_${labelId}_SiteId"><vt:lpwstr>00000000-0000-0000-0000-0000000000aa</vt:lpwstr></property></Properties>`;
  const name = Buffer.from("docProps/custom.xml");
  const data = Buffer.from(custom);
  const local = Buffer.alloc(30);
  local.writeUInt32LE(0x04034b50, 0);
  local.writeUInt16LE(20, 4);
  local.writeUInt32LE(data.length, 18);
  local.writeUInt32LE(data.length, 22);
  local.writeUInt16LE(name.length, 26);
  const central = Buffer.alloc(46);
  central.writeUInt32LE(0x02014b50, 0);
  central.writeUInt16LE(20, 6);
  central.writeUInt32LE(data.length, 20);
  central.writeUInt32LE(data.length, 24);
  central.writeUInt16LE(name.length, 28);
  central.writeUInt32LE(0, 42);
  const dirOffset = local.length + name.length + data.length;
  const dirSize = central.length + name.length;
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(1, 8);
  end.writeUInt16LE(1, 10);
  end.writeUInt32LE(dirSize, 12);
  end.writeUInt32LE(dirOffset, 16);
  return Buffer.concat([local, name, data, central, name, end]).toString("base64");
}

const LABEL_IDS = { internal: "aaaaaaaa-0000-4000-8000-000000000005", confidential: "aaaaaaaa-0000-4000-8000-000000000006" };

const LABELS = [
  { id: "aaaaaaaa-0000-4000-8000-000000000001", name: "Öffentlich" },
  { id: "aaaaaaaa-0000-4000-8000-000000000002", name: "Vertraulich" },
];

test("Labels: a labelled file shows its label, and the mail gets the same sensitivity", async ({ page }) => {
  const content = labelledDocx(LABELS[1].id);
  await openTaskpane(page, {
    attachments: [{ id: "a1", name: "Angebot.docx", size: 2 * MiB, content }],
    labels: LABELS,
    mailLabel: LABELS[0].id,
  });
  const row = page.locator("#attachments .list-item").first();
  await expect(row.locator(".file-label")).toHaveText("Vertraulich");
  await expect(row.locator("input")).toHaveAttribute("aria-label", /Vertraulich$/);
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("„Angebot.docx“ ist als „Vertraulich“ gekennzeichnet – die Nachricht hat dieselbe Vertraulichkeit erhalten.");
  expect((await officeLog(page)).mailLabel).toBe(LABELS[1].id);
});

test("Labels: never lowers the mail's label", async ({ page }) => {
  await openTaskpane(page, {
    attachments: [{ id: "a1", name: "Info.docx", size: 2 * MiB, content: labelledDocx(LABELS[0].id) }],
    labels: LABELS,
    mailLabel: LABELS[1].id,
  });
  await expect(page.locator("#attachments .file-label")).toHaveText("Öffentlich");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  await expect(page.locator("#status")).not.toContainText("Vertraulichkeit");
  expect((await officeLog(page)).mailLabel).toBeUndefined();
});

test("Labels: a protected file shows a note below the list and needs a confirmation", async ({ page }) => {
  const calls = await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pfile", 2 * MiB)] });
  const row = page.locator("#attachments .list-item").first();
  await expect(row.locator(".file-label")).toHaveText("Geschützt");
  await expect(row.locator(".label-lock")).toHaveCount(1);
  // A note right below the list; inserting only after "Trotzdem als Link senden".
  await expect(page.locator("#validity-section")).toBeHidden(); // "Gültigkeit" stays as it was
  await expect(page.locator("#encrypted-text")).toHaveText(
    "„Vertrag.pfile“ ist verschlüsselt. Empfänger außerhalb Ihrer Organisation können sie meist nicht öffnen.");
  await expect(page.locator("#submit")).toBeDisabled();
  await expectAccessible(page);
  // Unticking the file takes the note away.
  await row.locator(".file-name").click(); // the row is a label: a click on the name toggles the box
  await expect(page.locator("#encrypted-field")).toBeHidden();
  await row.locator(".file-name").click();
  await expect(page.locator("#encrypted-field")).toBeVisible();
  expect(calls.some((c) => c.path === "/api/uploads")).toBe(false);


  await page.locator("#encrypted-confirm-row .checkbox-text").click();
  await expect(page.locator("#encrypted-confirm")).toBeChecked();
  await expect(page.locator("#submit")).toBeEnabled();
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("Der Link wurde eingefügt.");
  expect(calls.some((c) => c.path === "/api/uploads")).toBe(true);
});

test("Labels: the \"×\" on the note removes the encrypted file from the mail – and the note goes with it", async ({ page }) => {
  await openTaskpane(page, { attachments: [attachment("a1", "Vertrag.pfile", 2 * MiB), attachment("a2", "Plan.pdf", 2 * MiB)] });
  await expect(page.locator("#encrypted-field")).toBeVisible();
  await expect(page.locator("#encrypted-dismiss")).toHaveAttribute("aria-label", "Datei entfernen – nicht als Link senden");
  await page.click("#encrypted-dismiss");
  expect((await officeLog(page)).removed).toEqual(["a1"]); // only the encrypted one
  await expect(page.locator("#attachments .list-item")).toHaveCount(1);
  await expect(page.locator("#attachments .file-name")).toHaveText("Plan.pdf");
  await expect(page.locator("#encrypted-field")).toBeHidden();
  await expect(page.locator("#submit")).toBeEnabled();
});

test("Labels: without Outlook's catalog the names and priority come from the tenant (Graph)", async ({ page }) => {
  // Purview's own priority: "Vertraulich" (rank 5) above "Intern" (rank 2), whatever order the list has.
  const tenantLabels = [
    { id: LABEL_IDS.confidential, name: "Vertraulich", rank: 5 },
    { id: LABEL_IDS.internal, name: "Intern", rank: 2 },
  ];
  await openTaskpane(page, {
    attachments: [{ id: "a1", name: "Angebot.docx", size: 2 * MiB, content: labelledDocx(LABEL_IDS.confidential) }],
    tenantLabels,
    mailLabel: LABEL_IDS.internal,
  });
  await expect(page.locator("#attachments .file-label")).toHaveText("Vertraulich");
  await page.click("#submit");
  await expect(page.locator("#status")).toContainText("„Angebot.docx“ ist als „Vertraulich“ gekennzeichnet");
  expect((await officeLog(page)).mailLabel).toBe(LABEL_IDS.confidential);
});

test("history, proofs, settings and help are pages of the account menu: back with the arrow or Escape", async ({ page }) => {
  await openTaskpane(page, { config: { deliveryProof: true } });
  await expect(page.locator(".tabbar, [role=tab]")).toHaveCount(0); // no tabs at all – "Senden" is the view
  await page.click("#account-toggle");
  // "Verlauf" right above "Zustellnachweise".
  await expect(page.locator("#account-menu button:visible")).toHaveText(
    ["Nextcloud öffnen", "Verlauf", "Zustellnachweise", "Einstellungen", "Hilfe", "Verbindung trennen"]);
  await page.keyboard.press("Escape");
  for (const [entry, panel, back] of [["#shares-open", "#panel-shares", "#shares-close"], ["#settings-open", "#panel-settings", "#settings-close"],
    ["#proofs-open", "#panel-proofs", "#proofs-close"], ["#help-toggle", "#panel-help", "#help-close"]]) {
    await page.click("#account-toggle");
    await page.click(entry);
    await expect(page.locator(panel)).toBeVisible();
    await expect(page.locator("#panel-send")).toBeHidden();
    await expect(page.locator(back)).toHaveAttribute("aria-label", "Zurück");
    await page.click(back);
    await expect(page.locator(panel)).toBeHidden();
    await expect(page.locator("#panel-send")).toBeVisible();
  }
  await page.click("#account-toggle");
  await page.click("#settings-open");
  await page.keyboard.press("Escape");
  await expect(page.locator("#panel-settings")).toBeHidden();
  await expect(page.locator("#panel-send")).toBeVisible();
  await expect(page.locator("#account-toggle")).toBeFocused();
});

test("Zustellnachweise: not in the account menu without DELIVERY_PROOF", async ({ page }) => {
  await openTaskpane(page);
  await page.click("#account-toggle");
  await expect(page.locator("#proofs-open")).toBeHidden();
});

test("Zustellnachweise: every proof in the account menu, filtered and searched, the seal opens the PDF", async ({ page }) => {
  const proofs = [
    { id: "20261118-1646-CBMH", createdAt: Date.parse("2026-11-18T15:46:00Z"), subject: "Unterlagen Neubau", sentAt: Date.parse("2026-11-18T15:47:00Z"),
      recipients: ["erika@example.com"], files: ["Angebot.zip"], downloads: 1, confirmed: 1, done: true,
      state: "fetched", fetched: 1, endsOn: null, keptUntil: Date.parse("2029-11-18T12:00:00Z"), kind: "file", mailAvailable: true, passwordTo: ["erika@example.com"] },
    { id: "20261112-0931-XK7P", createdAt: Date.parse("2026-11-12T08:31:00Z"), subject: "Vertrag 2026", sentAt: Date.parse("2026-11-12T08:32:00Z"),
      recipients: ["kunde@example.org"], files: ["Vertrag.pdf"], downloads: 0, confirmed: 0, done: false,
      state: "sent", fetched: 0, endsOn: "2026-11-26", keptUntil: Date.parse("2029-11-12T12:00:00Z"), kind: "file", mailAvailable: false, passwordTo: ["erika@example.com"] },
  ];
  await openTaskpane(page, { config: { deliveryProof: true }, proofs });
  await page.click("#account-toggle");
  await page.click("#proofs-open");
  await expect(page.locator("#panel-proofs")).toBeVisible();
  await expect(page.locator("#proofs-search")).toBeHidden(); // behind the magnifier, like in "Verlauf"
  const entries = page.locator("#proofs-list .proof-entry");
  await expect(entries).toHaveCount(2);
  await expect(entries.nth(0).locator(".share-title")).toHaveText("Unterlagen Neubau");
  // Message and password to the same address: one line instead of the address twice.
  await expect(entries.nth(0).locator(".proof-sent")).toHaveText("18.11.2026 · Nachricht und Kennwort an erika@example.com");
  await expect(entries.nth(0).locator(".proof-password")).toBeHidden();
  // Like a row of "Verlauf": the file with its seal, to whom, the state with its dot, the tracking ID.
  await expect(entries.nth(0).locator(".share-item")).toHaveText("Angebot.zip");
  await expect(entries.nth(0).locator(".state-text")).toHaveText("1× heruntergeladen");
  await expect(entries.nth(0).locator(".share-valid")).toHaveText("Aufbewahrt bis 18.11.2029"); // a finished proof: how long it is kept
  await expect(entries.nth(0).locator(".share-state")).toHaveClass(/done/);
  // Under the file: only the tracking ID (no "Vorgang") and after it the seal that opens the PDF – not after the file.
  await expect(entries.nth(0).locator(".proof-case")).toHaveText("20261118-1646-CBMH");
  await expect(entries.nth(0).locator(".proof-case .proof-seal")).toHaveCount(1);
  await expect(entries.nth(0).locator(".share-item .proof-seal")).toHaveCount(0);
  // The tracking ID bold right under the file, then to whom, "Gültig bis" and the state; no icons on the right.
  await expect(entries.nth(0).locator(".proof-case")).toHaveCSS("font-weight", "600");
  await expect(entries.nth(0).locator(".share-item + .proof-case + .proof-sent")).toHaveCount(1);
  await expect(entries.nth(0).locator(".proof-password + .share-valid + .share-state")).toHaveCount(1);
  await expect(entries.nth(0).locator(".share-actions")).toHaveCount(0);
  await expect(entries.nth(1).locator(".state-text")).toHaveText("Noch nicht abgerufen");
  await expect(entries.nth(1).locator(".share-valid")).toHaveText("Gültig bis 26.11.2026");
  // The password went to another address: two lines.
  await expect(entries.nth(1).locator(".proof-sent")).toHaveText("12.11.2026 · Nachricht an kunde@example.org");
  await expect(entries.nth(1).locator(".proof-password")).toHaveText("Kennwort an erika@example.com");
  // The subject opens the sent mail, as in "Verlauf" – only where the mail is known.
  await entries.nth(0).locator("button.share-title").click();
  await expect.poll(async () => (await officeLog(page)).openedMails).toContain("ews-proof-20261118-1646-CBMH");
  await expect(entries.nth(1).locator("button.share-title")).toHaveCount(0);
  // No second PDF button: the seal after the file name opens the PDF and tells the state by its colour (and tooltip).
  await expect(entries.nth(0).locator(".share-actions")).toHaveCount(0);
  await expect(entries.nth(0).locator(".proof-seal")).toHaveClass(/state-fetched/);
  await expect(entries.nth(0).locator(".proof-seal .icon")).toHaveCSS("color", /.+/);
  expect(await entries.nth(0).locator(".proof-seal .icon").evaluate((el) => getComputedStyle(el).color))
    .not.toBe(await entries.nth(1).locator(".proof-seal .icon").evaluate((el) => getComputedStyle(el).color));
  await expect(page.locator("#proofs-export")).toHaveCount(0);
  await expect(entries.nth(1).locator(".proof-seal")).toHaveAttribute("title", "Zustellnachweis 20261112-0931-XK7P als PDF öffnen – gesendet, noch nicht abgerufen");
  await expectAccessible(page);

  // No filters any more – the search is the way to find a proof.
  await expect(page.locator("#proofs-filter")).toHaveCount(0);

  // Search (behind the magnifier): by recipient, by tracking ID (any case), nothing.
  await page.click("#proofs-search-toggle");
  await expect(page.locator("#proofs-search")).toBeFocused();
  await page.fill("#proofs-search", "kunde@");
  await expect(entries).toHaveCount(1);
  await expect(entries.nth(0).locator(".share-title")).toHaveText("Vertrag 2026");
  await page.fill("#proofs-search", "cbmh");
  await expect(entries).toHaveCount(1);
  await page.fill("#proofs-search", "nichts");
  await expect(entries).toHaveCount(0);
  await expect(page.locator("#proofs-empty")).toHaveText("Kein Zustellnachweis passt zur Suche.");

  await page.fill("#proofs-search", "");
  await entries.nth(1).locator(".proof-seal").click();
  await expect.poll(async () => (await officeLog(page)).openedUrls).toContain(
    "https://cloudshare.example.com/proof/20261112-0931-XK7P.pdf?until=1&lang=de&sig=x");
  await page.click("#proofs-close");
  await expect(page.locator("#panel-proofs")).toBeHidden();
});
