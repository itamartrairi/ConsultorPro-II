// IndexedDB Cache for Gemini AI Responses

const DB_NAME = 'sgf_ai_cache_db';
const STORE_NAME = 'ai_responses';
const DB_VERSION = 1;

interface CacheEntry {
  key: string;
  data: any;
  timestamp: number;
  expiresAt?: number;
}

// In-memory memory cache fallback for ultra-fast (0ms) synchronous access
const memoryCache = new Map<string, { data: any; timestamp: number }>();

function openDB(): Promise<IDBDatabase | null> {
  if (typeof window === 'undefined' || !window.indexedDB) {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    try {
      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = (event: any) => {
        const db = event.target.result;
        if (!db.objectStoreNames.contains(STORE_NAME)) {
          db.createObjectStore(STORE_NAME, { keyPath: 'key' });
        }
      };

      request.onsuccess = (event: any) => {
        resolve(event.target.result);
      };

      request.onerror = () => {
        console.warn("[AICache] IndexedDB unavailable, falling back to memory/localStorage");
        resolve(null);
      };
    } catch {
      resolve(null);
    }
  });
}

/**
 * Creates a deterministic string hash from any object or string
 */
export function generateAICacheKey(prefix: string, payload: any): string {
  try {
    const raw = typeof payload === 'string' ? payload : JSON.stringify(payload);
    let hash = 0;
    for (let i = 0; i < raw.length; i++) {
      const char = raw.charCodeAt(i);
      hash = (hash << 5) - hash + char;
      hash |= 0; // Convert to 32bit integer
    }
    return `${prefix}_${Math.abs(hash).toString(36)}_${raw.length}`;
  } catch {
    return `${prefix}_${Date.now()}`;
  }
}

/**
 * Get cached AI response from Memory -> IndexedDB -> LocalStorage
 */
export async function getCachedAI<T = any>(key: string): Promise<T | null> {
  // 1. Check memory cache (0ms lookup)
  if (memoryCache.has(key)) {
    const entry = memoryCache.get(key)!;
    return entry.data as T;
  }

  // 2. Check IndexedDB
  try {
    const db = await openDB();
    if (db) {
      const tx = db.transaction(STORE_NAME, 'readonly');
      const store = tx.objectStore(STORE_NAME);
      const req = store.get(key);

      const dbResult: CacheEntry | undefined = await new Promise((resolve) => {
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(undefined);
      });

      if (dbResult && dbResult.data !== undefined) {
        // Populate memory cache
        memoryCache.set(key, { data: dbResult.data, timestamp: dbResult.timestamp });
        return dbResult.data as T;
      }
    }
  } catch (e) {
    console.warn("[AICache] Error reading from IndexedDB:", e);
  }

  // 3. Check localStorage fallback
  try {
    const lsItem = localStorage.getItem(`ai_cache_${key}`);
    if (lsItem) {
      const parsed = JSON.parse(lsItem);
      memoryCache.set(key, { data: parsed.data, timestamp: parsed.timestamp });
      return parsed.data as T;
    }
  } catch {}

  return null;
}

/**
 * Save AI response to Memory, IndexedDB and LocalStorage
 */
export async function setCachedAI(key: string, data: any, ttlDays = 30): Promise<void> {
  if (data === null || data === undefined) return;

  const now = Date.now();
  const expiresAt = now + ttlDays * 24 * 60 * 60 * 1000;

  // 1. Save to memory cache
  memoryCache.set(key, { data, timestamp: now });

  // 2. Save to IndexedDB
  try {
    const db = await openDB();
    if (db) {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.put({
        key,
        data,
        timestamp: now,
        expiresAt
      });
    }
  } catch (e) {
    console.warn("[AICache] Error saving to IndexedDB:", e);
  }

  // 3. Save small/critical responses to localStorage as backup
  try {
    const serialized = JSON.stringify({ data, timestamp: now });
    if (serialized.length < 50000) {
      localStorage.setItem(`ai_cache_${key}`, serialized);
    }
  } catch {}
}

/**
 * Clear all cached AI entries
 */
export async function clearAICache(): Promise<void> {
  memoryCache.clear();
  try {
    const db = await openDB();
    if (db) {
      const tx = db.transaction(STORE_NAME, 'readwrite');
      const store = tx.objectStore(STORE_NAME);
      store.clear();
    }
  } catch {}

  try {
    for (let i = 0; i < localStorage.length; i++) {
      const k = localStorage.key(i);
      if (k && k.startsWith('ai_cache_')) {
        localStorage.removeItem(k);
      }
    }
  } catch {}
}
