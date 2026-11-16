import { existsSync, readFileSync, statSync } from "node:fs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
import { describe, expect, it } from "vitest";
import { getCredentials, removeCredentials, saveCredentials } from "../src/credentials.js";
import { PersistentMap } from "../src/db.js";
import { seal, unseal } from "../src/storage.js";

const dbFile = () => path.join(process.env.DATA_DIR!, "cloudshare.db");
/** The rows as they are on disk, read with a second connection (what a restart would load). */
const rows = (store: string) =>
  new DatabaseSync(dbFile(), { readOnly: true }).prepare("SELECT key, data FROM entries WHERE store = ?").all(store) as Array<{
    key: string;
    data: string;
  }>;

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

describe("database (SQLite)", () => {
  it("writes one row per entry, sealed, and loads them again", () => {
    const map = new PersistentMap<{ name: string; count: number }>("test-store");
    map.set("a", { name: "Präsentation.pptx", count: 1 });
    map.set("b", { name: "Angebot.pdf", count: 2 });
    map.get("a")!.count = 5;
    map.save("a"); // changed in place
    map.delete("b");

    expect(rows("test-store").map((r) => r.key)).toEqual(["a"]);
    expect(rows("test-store")[0].data).not.toContain("Präsentation"); // nothing readable on disk

    const reloaded = new PersistentMap<{ name: string; count: number }>("test-store");
    expect([...reloaded.entries()]).toEqual([["a", { name: "Präsentation.pptx", count: 5 }]]);
    expect(new PersistentMap("other-store").size).toBe(0); // stores are separate
  });

  it("is readable only by the CloudShare user", () => {
    new PersistentMap("mode-check").set("x", 1);
    expect(statSync(dbFile()).mode & 0o777).toBe(0o600);
  });
});

describe("credentials", () => {
  it("stores app passwords encrypted and bound to the user", async () => {
    await saveCredentials("oid-1", { loginName: "max", appPassword: "app-pass-123" });
    expect(getCredentials("oid-1")).toEqual({ loginName: "max", appPassword: "app-pass-123" });

    // Neither the database nor its write-ahead log hold the password in plain text.
    const onDisk = [dbFile(), `${dbFile()}-wal`].filter(existsSync).map((f) => readFileSync(f).toString("latin1")).join("");
    expect(onDisk).not.toContain("app-pass-123");
    expect(rows("credentials").map((r) => r.key)).toContain("oid-1");
    expect(getCredentials("oid-2")).toBeUndefined();

    await removeCredentials("oid-1");
    expect(getCredentials("oid-1")).toBeUndefined();
    expect(rows("credentials").map((r) => r.key)).not.toContain("oid-1");
  });
});
