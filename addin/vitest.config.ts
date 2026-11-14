import { defineConfig } from "vitest/config";

// Tests cover the pure modules (password, footer, mail card) and the event handlers with an Office.js mock.
export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    // German texts from locales/de.json (in the container the server injects them into the bundles).
    setupFiles: ["./test/setup.ts"],
  },
});
