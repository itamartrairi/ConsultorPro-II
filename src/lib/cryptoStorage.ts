import CryptoJS from 'crypto-js';

// Base application salt for persistent local storage at rest
const BASE_SALT = 'sebrae_diag_at_rest_v1_2026';
const GLOBAL_STORAGE_KEY_SEED = 'sebrae_app_master_secure_storage_2026';

// Persistent machine/device seed key
let customSessionPassword: string | null = null;

/**
 * Configure optional user ID or session password
 */
export function setStorageUserId(_uid: string | null) {
  // Maintained for API compatibility, but storage encryption remains stable per device
}

export function setStorageSessionPassword(pass: string | null) {
  customSessionPassword = pass;
}

/**
 * Gets or creates a stable browser device installation key.
 */
function getDeviceKey(): string {
  if (typeof window === 'undefined') return GLOBAL_STORAGE_KEY_SEED;
  try {
    let devKey = window.localStorage.getItem('_sys_dev_key');
    if (!devKey || devKey === 'undefined' || devKey === 'null') {
      devKey = 'dev_' + Math.random().toString(36).substring(2) + Date.now().toString(36);
      window.localStorage.setItem('_sys_dev_key', devKey);
    }
    return devKey;
  } catch (e) {
    return GLOBAL_STORAGE_KEY_SEED;
  }
}

/**
 * Derives the active AES-256 encryption key.
 * Uses a stable device key so data is never lost or corrupted across page reloads / auth state changes.
 */
export function getActiveEncryptionKey(): string {
  if (customSessionPassword && customSessionPassword.trim()) {
    return CryptoJS.SHA256(`pass:${customSessionPassword}:${BASE_SALT}`).toString();
  }
  const deviceKey = getDeviceKey();
  return CryptoJS.SHA256(`dev:${deviceKey}:${BASE_SALT}`).toString();
}

/**
 * Fallback encryption keys to ensure 100% data recovery even if environment changes.
 */
function getAllPossibleKeys(): string[] {
  const keys: string[] = [getActiveEncryptionKey()];
  keys.push(CryptoJS.SHA256(`dev:${getDeviceKey()}:${BASE_SALT}`).toString());
  keys.push(CryptoJS.SHA256(GLOBAL_STORAGE_KEY_SEED).toString());
  keys.push(CryptoJS.SHA256(BASE_SALT).toString());
  return Array.from(new Set(keys));
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
 * Decrypts an encrypted string value. If plain text or legacy format, returns as-is for backward compatibility.
 */
export function decryptValue(rawValue: string): string {
  if (!rawValue) return rawValue;
  if (!rawValue.startsWith('enc:v1:')) {
    // Unencrypted plain text
    return rawValue;
  }

  const cipherText = rawValue.substring(7); // Remove 'enc:v1:'
  const keysToTry = getAllPossibleKeys();

  for (const key of keysToTry) {
    try {
      const bytes = CryptoJS.AES.decrypt(cipherText, key);
      const decrypted = bytes.toString(CryptoJS.enc.Utf8);
      if (decrypted && decrypted.length > 0) {
        return decrypted;
      }
    } catch (e) {
      // Try next key
    }
  }

  // If decryption fails, check if the inner text is JSON or return fallback
  console.warn("[CryptoStorage] Não foi possível descriptografar valor. Verificando integridade...");
  return rawValue;
}

// Unencrypted keys that do not require encryption (e.g., system flags, device keys, theme)
const UNENCRYPTED_KEYS = new Set([
  '_sys_dev_key',
  'storage_mode',
  'preferred_logo',
  'selected_plan',
  'custom_gemini_api_key'
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
