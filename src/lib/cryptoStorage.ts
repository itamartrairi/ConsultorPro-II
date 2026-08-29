import CryptoJS from 'crypto-js';

// Base application salt for key derivation
const BASE_SALT = 'sebrae_diag_at_rest_v1_2026';

// Persistent machine/device seed key in memory or unencrypted metadata
let currentUserId: string | null = null;
let customSessionPassword: string | null = null;

/**
 * Configure active User ID or session password for local storage encryption key derivation.
 */
export function setStorageUserId(uid: string | null) {
  currentUserId = uid;
}

export function setStorageSessionPassword(pass: string | null) {
  customSessionPassword = pass;
}

/**
 * Gets or creates a browser device installation key.
 */
function getDeviceKey(): string {
  if (typeof window === 'undefined') return 'server_fallback_key';
  try {
    let devKey = window.localStorage.getItem('_sys_dev_key');
    if (!devKey) {
      devKey = 'dev_' + Math.random().toString(36).substring(2) + Date.now().toString(36);
      window.localStorage.setItem('_sys_dev_key', devKey);
    }
    return devKey;
  } catch (e) {
    return 'mem_device_key';
  }
}

/**
 * Derives the active AES-256 encryption key.
 */
export function getActiveEncryptionKey(): string {
  if (customSessionPassword && customSessionPassword.trim()) {
    return CryptoJS.SHA256(`pass:${customSessionPassword}:${BASE_SALT}`).toString();
  }
  if (currentUserId && currentUserId.trim()) {
    return CryptoJS.SHA256(`uid:${currentUserId}:${BASE_SALT}`).toString();
  }
  const deviceKey = getDeviceKey();
  return CryptoJS.SHA256(`dev:${deviceKey}:${BASE_SALT}`).toString();
}

/**
 * Encrypts a string value using AES-256.
 */
export function encryptValue(plainText: string): string {
  if (!plainText) return plainText;
  try {
    const key = getActiveEncryptionKey();
    const cipherText = CryptoJS.AES.encrypt(plainText, key).toString();
    return `enc:v1:${cipherText}`;
  } catch (e) {
    console.error("Erro na criptografia dos dados:", e);
    return plainText;
  }
}

/**
 * Decrypts an encrypted string value. If plain text, returns as-is for backward compatibility.
 */
export function decryptValue(rawValue: string): string {
  if (!rawValue) return rawValue;
  if (!rawValue.startsWith('enc:v1:')) {
    // Unencrypted legacy plain text
    return rawValue;
  }

  const cipherText = rawValue.substring(7); // Remove 'enc:v1:'
  try {
    const key = getActiveEncryptionKey();
    const bytes = CryptoJS.AES.decrypt(cipherText, key);
    const decrypted = bytes.toString(CryptoJS.enc.Utf8);
    
    if (decrypted && decrypted.length > 0) {
      return decrypted;
    }
  } catch (e) {
    console.warn("Falha ao descriptografar com chave primária, tentando fallback device key...");
  }

  // Fallback to device key if user key changed
  try {
    const fallbackKey = CryptoJS.SHA256(`dev:${getDeviceKey()}:${BASE_SALT}`).toString();
    const bytes = CryptoJS.AES.decrypt(cipherText, fallbackKey);
    const decrypted = bytes.toString(CryptoJS.enc.Utf8);
    if (decrypted && decrypted.length > 0) {
      return decrypted;
    }
  } catch (e) {
    // Ignore fallback failure
  }

  // If all fails, return raw string
  return rawValue;
}

// Unencrypted keys that do not require encryption (e.g., system flags, device keys)
const UNENCRYPTED_KEYS = new Set([
  '_sys_dev_key',
  'storage_mode',
  'preferred_logo'
]);

/**
 * Safe local storage wrapper with automatic AES-256 Encryption at Rest.
 */
const memoryStorage: Record<string, string> = {};

export const encryptedLocalStorage = {
  getItem: (key: string): string | null => {
    let rawVal: string | null = null;
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        rawVal = window.localStorage.getItem(key);
      } else {
        rawVal = memoryStorage[key] || null;
      }
    } catch (e) {
      rawVal = memoryStorage[key] || null;
    }

    if (!rawVal) return null;
    if (UNENCRYPTED_KEYS.has(key)) return rawVal;

    return decryptValue(rawVal);
  },

  setItem: (key: string, value: string): void => {
    if (value === null || value === undefined) {
      encryptedLocalStorage.removeItem(key);
      return;
    }

    const valueToStore = UNENCRYPTED_KEYS.has(key) ? value : encryptValue(value);

    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.setItem(key, valueToStore);
      } else {
        memoryStorage[key] = valueToStore;
      }
    } catch (e) {
      memoryStorage[key] = valueToStore;
    }
  },

  removeItem: (key: string): void => {
    try {
      if (typeof window !== 'undefined' && window.localStorage) {
        window.localStorage.removeItem(key);
      } else {
        delete memoryStorage[key];
      }
    } catch (e) {
      delete memoryStorage[key];
    }
  },

  /**
   * Migrate existing unencrypted plain-text entries in localStorage to enc:v1:
   */
  migrateAllToEncrypted: (): number => {
    if (typeof window === 'undefined' || !window.localStorage) return 0;
    let count = 0;
    try {
      const keysToMigrate: string[] = [];
      for (let i = 0; i < window.localStorage.length; i++) {
        const k = window.localStorage.key(i);
        if (k && !UNENCRYPTED_KEYS.has(k)) {
          keysToMigrate.push(k);
        }
      }

      for (const k of keysToMigrate) {
        const val = window.localStorage.getItem(k);
        if (val && !val.startsWith('enc:v1:')) {
          const enc = encryptValue(val);
          window.localStorage.setItem(k, enc);
          count++;
        }
      }
    } catch (e) {
      console.error("Erro durante a migração para criptografia local:", e);
    }
    return count;
  }
};
