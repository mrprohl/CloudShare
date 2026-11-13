import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { getCredentials, removeCredentials, saveCredentials } from "../src/credentials.js";
import { createAtomicWriter, seal, unseal } from "../src/storage.js";

describe("seal / unseal", () => {
  it("round-trips text", () => {
    const sealed = seal("geheim – äöü", "context");
    expect(unseal(sealed, "context")).toBe("geheim – äöü");
  });

  it("uses a fresh IV for every call", () => {
    expect(seal("x", "a").iv).not.toBe(seal("x", "a").iv);
  });

  it("refuses data sealed for another context (AAD) or tampered with", () => {
    const sealed = seal("geheim", "user-a");
    expect(() => unseal(sealed, "user-b")).toThrow();
    const tampered = { ...sealed, data: Buffer.from("anders").toString("base64") };
    expect(() => unseal(tampered, "user-a")).toThrow();
  });
});

describe("createAtomicWriter", () => {
  it("writes the last content when called concurrently", async () => {
    const file = path.join(process.env.DATA_DIR!, "writer.json");
    const write = createAtomicWriter(file);
    await Promise.all([write("1"), write("2"), write("3")]);
    expect(readFileSync(file, "utf8")).toBe("3");
  });
});

describe("credentials", () => {
  it("stores app passwords encrypted and bound to the user", async () => {
    await saveCredentials("oid-1", { loginName: "max", appPassword: "app-pass-123" });
    expect(getCredentials("oid-1")).toEqual({ loginName: "max", appPassword: "app-pass-123" });

    const raw = readFileSync(path.join(process.env.DATA_DIR!, "nextcloud-credentials.json"), "utf8");
    expect(raw).not.toContain("app-pass-123");
    expect(getCredentials("oid-2")).toBeUndefined();

    await removeCredentials("oid-1");
    expect(getCredentials("oid-1")).toBeUndefined();
  });
});
