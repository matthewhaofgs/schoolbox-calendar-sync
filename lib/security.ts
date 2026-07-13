export class HttpError extends Error {
  constructor(
    public readonly status: number,
    message: string,
    public readonly detail?: string,
  ) {
    super(message);
    this.name = "HttpError";
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) bytes[index] = binary.charCodeAt(index);
  return bytes;
}

async function encryptionKey(): Promise<CryptoKey> {
  const configured = process.env.CONFIG_ENCRYPTION_KEY?.trim();
  if (!configured) {
    throw new HttpError(
      503,
      "Credential encryption is not configured",
      "Set CONFIG_ENCRYPTION_KEY to 32 random bytes encoded as base64 before saving credentials.",
    );
  }

  let keyBytes: Uint8Array;
  try {
    keyBytes = base64ToBytes(configured);
  } catch {
    throw new HttpError(503, "CONFIG_ENCRYPTION_KEY is not valid base64");
  }
  if (keyBytes.byteLength !== 32) {
    throw new HttpError(503, "CONFIG_ENCRYPTION_KEY must decode to exactly 32 bytes");
  }

  return crypto.subtle.importKey("raw", keyBytes as BufferSource, "AES-GCM", false, ["encrypt", "decrypt"]);
}

export async function encryptSecret(value: string): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const encoded = new TextEncoder().encode(value);
  const ciphertext = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, await encryptionKey(), encoded);
  return `v1.${bytesToBase64(iv)}.${bytesToBase64(new Uint8Array(ciphertext))}`;
}

export async function decryptSecret(value: string): Promise<string> {
  const [version, ivPart, ciphertextPart] = value.split(".");
  if (version !== "v1" || !ivPart || !ciphertextPart) throw new Error("Unsupported encrypted credential format");
  const iv = base64ToBytes(ivPart);
  const ciphertext = base64ToBytes(ciphertextPart);
  const plaintext = await crypto.subtle.decrypt(
    { name: "AES-GCM", iv: iv as BufferSource },
    await encryptionKey(),
    ciphertext as BufferSource,
  );
  return new TextDecoder().decode(plaintext);
}

export function jsonError(error: unknown): Response {
  if (error instanceof HttpError) {
    return Response.json({ error: error.message, detail: error.detail }, { status: error.status });
  }
  const message = error instanceof Error ? error.message : "Unexpected error";
  return Response.json({ error: message }, { status: 500 });
}
