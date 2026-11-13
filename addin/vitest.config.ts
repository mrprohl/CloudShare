import { defineConfig } from "vitest/config";

// Tests cover the pure modules (password, footer, mail card); Office.js-dependent code is exercised in Outlook.
export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
  },
});
