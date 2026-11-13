import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    environment: "node",
    include: ["test/**/*.test.ts"],
    // Sets the environment variables that src/config.ts requires before any module is imported.
    setupFiles: ["./test/setup.ts"],
    // Each test file gets its own module instances (jobs/credentials keep state in module scope).
    isolate: true,
  },
});
