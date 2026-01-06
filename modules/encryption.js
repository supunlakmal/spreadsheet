/**
 * Encryption Module for Spreadsheet Application
 * Provides AES-GCM 256-bit encryption with PBKDF2 key derivation
 */

// ========== Encryption Module (AES-GCM 256-bit) ==========
export const CryptoUtils = {
  algo: { name: "AES-GCM", length: 256 },
  kdf: { name: "PBKDF2", hash: "SHA-256", iterations: 100000 },

  // Derive a cryptographic key from a password using PBKDF2
  async deriveKey(password, salt) {
    const enc = new TextEncoder();
    const keyMaterial = await window.crypto.subtle.importKey("raw", enc.encode(password), "PBKDF2", false, ["deriveKey"]);
    return window.crypto.subtle.deriveKey({ ...this.kdf, salt: salt }, keyMaterial, this.algo, false, ["encrypt", "decrypt"]);
  },

  // Encrypt data string with password, returns Base64 string
  async encrypt(dataString, password) {
    const enc = new TextEncoder();
    const salt = window.crypto.getRandomValues(new Uint8Array(16));
    const iv = window.crypto.getRandomValues(new Uint8Array(12));

    const key = await this.deriveKey(password, salt);
    const encodedData = enc.encode(dataString);

    const encryptedContent = await window.crypto.subtle.encrypt({ name: "AES-GCM", iv: iv }, key, encodedData);

    // Pack: Salt (16) + IV (12) + EncryptedData
    const buffer = new Uint8Array(salt.byteLength + iv.byteLength + encryptedContent.byteLength);
    buffer.set(salt, 0);
    buffer.set(iv, salt.byteLength);
    buffer.set(new Uint8Array(encryptedContent), salt.byteLength + iv.byteLength);

    return this.bufferToBase64(buffer);
  },

  // Decrypt Base64 string with password, returns original data string
  async decrypt(base64String, password) {
    const buffer = this.base64ToBuffer(base64String);

    // Extract: Salt (16) + IV (12) + EncryptedData
    const salt = buffer.slice(0, 16);
    const iv = buffer.slice(16, 28);
    const data = buffer.slice(28);

    const key = await this.deriveKey(password, salt);

    const decryptedContent = await window.crypto.subtle.decrypt({ name: "AES-GCM", iv: iv }, key, data);

    const dec = new TextDecoder();
    return dec.decode(decryptedContent);
  },

  // Convert Uint8Array to URL-safe Base64
  bufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    for (let i = 0; i < bytes.byteLength; i++) {
      binary += String.fromCharCode(bytes[i]);
    }
    // Use URL-safe Base64 (replace + with -, / with _, remove padding)
    return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  },

  // Convert URL-safe Base64 to Uint8Array
  base64ToBuffer(base64) {
    // Restore standard Base64
    let standardBase64 = base64.replace(/-/g, "+").replace(/_/g, "/");
    // Add padding if needed
    while (standardBase64.length % 4) {
      standardBase64 += "=";
    }
    const binaryString = atob(standardBase64);
    const bytes = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes;
  },
};

// ========== Encryption Codec (Wrapper for CryptoUtils) ==========
// Handles encryption prefix and provides clean interface for encrypt/decrypt
export const EncryptionCodec = {
  PREFIX: "ENC:",

  // Check if a URL hash represents encrypted data
  isEncrypted(hash) {
    return hash && hash.startsWith(this.PREFIX);
  },

  // Wrap encrypted data with the encryption prefix
  wrap(encryptedData) {
    return this.PREFIX + encryptedData;
  },

  // Unwrap the prefix from encrypted data
  unwrap(prefixedData) {
    return prefixedData.slice(this.PREFIX.length);
  },

  // Encrypt a payload string and wrap with prefix
  async encrypt(payload, password) {
    const encrypted = await CryptoUtils.encrypt(payload, password);
    return this.wrap(encrypted);
  },

  // Decrypt data (handles both wrapped and unwrapped formats)
  async decrypt(wrappedData, password) {
    const data = this.isEncrypted(wrappedData) ? this.unwrap(wrappedData) : wrappedData;
    return await CryptoUtils.decrypt(data, password);
  },
};
