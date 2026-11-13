import { defineConfig, devices } from "@playwright/test";

// UI tests of the taskpane in Chromium at taskpane width; Office.js and the backend API are mocked.
// Run: npm run test:ui (builds the add-in first). Report and screenshots: addin/playwright-report.
export default defineConfig({
  testDir: ".",
  testMatch: "*.e2e.ts",
  fullyParallel: true,
  forbidOnly: !!process.env.CI,
  retries: process.env.CI ? 1 : 0,
  reporter: [["list"], ["html", { outputFolder: "../playwright-report", open: "never" }]],
  outputDir: "../test-results",
  use: {
    ...devices["Desktop Chrome"],
    baseURL: "http://127.0.0.1:4173",
    viewport: { width: 376, height: 900 }, // width of the Outlook taskpane
    locale: "de-DE",
    timezoneId: "Europe/Berlin",
    trace: "retain-on-failure",
  },
  webServer: {
    command: "node serve.mjs", // cwd = directory of this config
    url: "http://127.0.0.1:4173/taskpane.html",
    reuseExistingServer: !process.env.CI,
  },
});
