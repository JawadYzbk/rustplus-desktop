/**
 * Secret-at-rest codec seam.
 *
 * Production wires Electron's safeStorage (DPAPI-backed on Windows). Tests and non-Electron contexts wire
 * a passthrough codec. The legacy app stored PlayerToken/webhook URLs in PLAINTEXT — encrypting them is a
 * deliberate, documented improvement (audit DATA_STORES §3/§6), not a format we must stay readable for.
 */
export interface SecretCodec {
  /** Encrypt a plaintext secret into an at-rest blob. */
  seal(plaintext: string): string;
  /** Recover the plaintext; throws if the blob was sealed by another machine/user profile. */
  open(blob: string): string;
}

export const SEALED_PREFIX = "enc:v1:";

/** Dev/test codec — identity. NEVER wire this in production composition roots. */
export class PassthroughSecretCodec implements SecretCodec {
  seal(plaintext: string): string {
    return plaintext;
  }
  open(blob: string): string {
    return blob;
  }
}
