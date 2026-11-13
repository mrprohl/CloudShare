/**
 * Shared helpers for the persistent files in DATA_DIR: AES-256-GCM sealing with CREDENTIAL_KEY and
 * atomic, sequential writes.
 */
import { createCipheriv, createDecipheriv, randomBytes } from "node:crypto";
import { mkdirSync } from "node:fs";
import { rename, writeFile } from "node:fs/promises";
import { config } from "./config.js";

const key = Buffer.from(config.credentialKey, "base64");
if (key.length !== 32) throw new Error("CREDENTIAL_KEY must be 32 bytes long (Base64)");

mkdirSync(config.dataDir, { recursive: true });

export interface Sealed {
  iv: string;
  tag: string;
  data: string;
}

/** Encrypts text; the AAD binds the ciphertext to its context (e.g. the user's oid). */
export function seal(plaintext: string, aad: string): Sealed {
  const iv = randomBytes(12);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  cipher.setAAD(Buffer.from(aad));
  const data = Buffer.concat([cipher.update(plaintext, "utf8"), cipher.final()]);
  return { iv: iv.toString("base64"), tag: cipher.getAuthTag().toString("base64"), data: data.toString("base64") };
}

/** Decrypts text sealed with the same key and AAD; throws if the key changed or the data was tampered with. */
export function unseal(sealed: Sealed, aad: string): string {
  const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(sealed.iv, "base64"));
  decipher.setAAD(Buffer.from(aad));
  decipher.setAuthTag(Buffer.from(sealed.tag, "base64"));
  return Buffer.concat([decipher.update(Buffer.from(sealed.data, "base64")), decipher.final()]).toString("utf8");
}

/** Returns a writer that replaces the file atomically (tmp + rename) and serialises concurrent writes. */
export function createAtomicWriter(file: string): (content: string) => Promise<void> {
  let queue: Promise<void> = Promise.resolve();
  return (content: string) => {
    queue = queue
      .catch(() => undefined)
      .then(async () => {
        const tmp = `${file}.tmp`;
        await writeFile(tmp, content, { mode: 0o600 });
        await rename(tmp, file);
      });
    return queue;
  };
}
