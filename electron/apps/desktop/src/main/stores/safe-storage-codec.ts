/**
 * SafeStorage-backed SecretCodec — the production composition.
 *
 * Blob format: "enc:v1:" + base64(safeStorage.encryptString(plaintext)). safeStorage on Windows uses
 * DPAPI (user+machine scoped), so blobs are non-portable by design; cross-machine restore goes through
 * the backup flow, not file copies.
 */
import { safeStorage } from "electron";
import { SEALED_PREFIX, type SecretCodec } from "./secret-codec.js";

export class SafeStorageSecretCodec implements SecretCodec {
  seal(plaintext: string): string {
    if (!plaintext) return plaintext;
    if (!safeStorage.isEncryptionAvailable()) {
      // Legacy behavior was plaintext-at-rest; keep the app working but say it loudly in the log
      // (composition root logs via returned marker; store layer surfaces this through its logger).
      throw new Error("safeStorage unavailable — refusing to write plaintext secrets (legacy parity would be insecure)");
    }
    return SEALED_PREFIX + safeStorage.encryptString(plaintext).toString("base64");
  }

  open(blob: string): string {
    if (!blob.startsWith(SEALED_PREFIX)) return blob; // legacy/plaintext value read back as-is
    if (!safeStorage.isEncryptionAvailable()) {
      throw new Error("safeStorage unavailable — cannot decrypt stored secret");
    }
    const raw = Buffer.from(blob.slice(SEALED_PREFIX.length), "base64");
    return safeStorage.decryptString(raw);
  }
}
