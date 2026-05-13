import { defineConfig, devices } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/e2e",
  timeout: 60_000,
  expect: {
    timeout: 10_000
  },
  use: {
    baseURL: "http://localhost:3010",
    trace: "retain-on-failure"
  },
  webServer: {
    command: "SYNTAX_AGENT_ENDPOINT=http://127.0.0.1:4011/invoke npm run dev:local",
    url: "http://localhost:3010",
    reuseExistingServer: false,
    timeout: 120_000
  },
  projects: [
    {
      name: "chromium",
      use: { ...devices["Desktop Chrome"] }
    }
  ]
});
