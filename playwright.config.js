import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests",
  testMatch: "**/e2e-playwright.spec.js",
  timeout: 30000,
  use: {
    headless: true,
    viewport: { width: 380, height: 700 }, // Word taskpane standard dimensions
    ignoreHTTPSErrors: true,
  },
  webServer: {
    command: "node server/server.js",
    port: 3011,
    reuseExistingServer: !process.env.CI,
    env: {
      PORT: "3011",
      LICENSE_REQUIRED: "false", // Allow tests to run without Supabase license
    },
  },
});
