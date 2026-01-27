import { decryptString, encryptString } from "./cryptoManager.js";

const HASH_PREFIX = "ENC:";

function ensureLzString() {
  if (!window.LZString) {
    throw new Error("LZString missing");
  }
  return window.LZString;
}

export function isEncryptedHash(hash = window.location.hash) {
  const raw = hash.startsWith("#") ? hash.slice(1) : hash;
  return raw.startsWith(HASH_PREFIX);
}

export async function readStateFromHash(password = null) {
  const raw = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : window.location.hash;
  if (!raw) return null;

  const lz = ensureLzString();
  if (raw.startsWith(HASH_PREFIX)) {
    if (!password) {
      const error = new Error("Password required");
      error.code = "PASSWORD_REQUIRED";
      throw error;
    }
    const payload = raw.slice(HASH_PREFIX.length);
    const compressed = await decryptString(payload, password);
    const json = lz.decompressFromEncodedURIComponent(compressed);
    if (!json) throw new Error("Decompression failed");
    return JSON.parse(json);
  }

  const json = lz.decompressFromEncodedURIComponent(raw);
  if (!json) throw new Error("Decompression failed");
  return JSON.parse(json);
}

export async function writeStateToHash(state, password = null) {
  const lz = ensureLzString();
  const json = JSON.stringify(state);
  const compressed = lz.compressToEncodedURIComponent(json);
  let nextHash = compressed;

  if (password) {
    const encrypted = await encryptString(compressed, password);
    nextHash = `${HASH_PREFIX}${encrypted}`;
  }

  if (window.location.hash !== `#${nextHash}`) {
    history.replaceState(null, "", `#${nextHash}`);
  }
  return nextHash;
}

export function clearHash() {
  history.replaceState(null, "", window.location.pathname);
}
