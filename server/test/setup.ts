/**
 * Test environment: dummy configuration (no real Entra, Nextcloud or Vaultwarden is contacted – every network call
 * in the tests goes to a mocked fetch) and a fresh temporary DATA_DIR per test file.
 */
import { mkdtempSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

Object.assign(process.env, {
  PUBLIC_BASE_URL: "https://cloudshare.example.test",
  ENTRA_TENANT_ID: "00000000-0000-0000-0000-000000000001",
  ENTRA_CLIENT_ID: "00000000-0000-0000-0000-000000000002",
  ENTRA_CLIENT_SECRET: "test-secret",
  NEXTCLOUD_URL: "https://cloud.example.test",
  // 32 bytes, Base64 – test key only
  CREDENTIAL_KEY: Buffer.alloc(32, 7).toString("base64"),
  BW_PASSWORD: "test-master-password",
  // SMS and download confirmation switched on, so their code paths are tested (seven.io and Graph are mocked).
  SEVEN_API_KEY: "test-seven-key",
  SEVEN_WEBHOOK_SECRET: "test-webhook-secret",
  DOWNLOAD_NOTICE: "true",
  TEMPLATES_FILE: path.join(tmpdir(), "cloudshare-no-templates.json"),
  DATA_DIR: mkdtempSync(path.join(tmpdir(), "cloudshare-test-")),
});
