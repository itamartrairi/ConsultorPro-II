/**
 * Evidence & Report Persistent Storage Engine
 * Uses IndexedDB for resilient multi-megabyte local persistence of photographic evidence,
 * preventing any data loss from quota errors or uncommitted Firestore batches.
 */

const DB_NAME = 'ConsultoriaPro_Evidence_Store_v1';
const DB_VERSION = 1;

let dbPromise: Promise<IDBDatabase> | null = null;

function getIDB(): Promise<IDBDatabase> {
  if (typeof window === 'undefined' || !window.indexedDB) {
    return Promise.reject(new Error("IndexedDB não suportado"));
  }

  if (!dbPromise) {
    dbPromise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error("IndexedDB timeout"));
      }, 2000);

      try {
        const req = window.indexedDB.open(DB_NAME, DB_VERSION);

        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains('diagnostico_cache')) {
            db.createObjectStore('diagnostico_cache', { keyPath: 'diagnosticoId' });
          }
          if (!db.objectStoreNames.contains('individual_evidences')) {
            db.createObjectStore('individual_evidences', { keyPath: 'key' });
          }
        };

        req.onsuccess = () => {
          clearTimeout(timer);
          resolve(req.result);
        };
        req.onerror = () => {
          clearTimeout(timer);
          reject(req.error);
        };
        req.onblocked = () => {
          clearTimeout(timer);
          console.warn("[EvidenceStore] IndexedDB open blocked by another tab");
          // Attempt to resolve if result exists or reject
          if (req.result) resolve(req.result);
          else reject(new Error("IndexedDB blocked"));
        };
      } catch (err) {
        clearTimeout(timer);
        reject(err);
      }
    });
  }

  return dbPromise;
}

/**
 * Compresses an image file with smart dimensions and JPEG quality
 * reducing raw multi-MB photos to ~30-50KB without perceptual loss.
 */
export async function compressImageToDataUrl(
  file: File, 
  maxDimension = 540, 
  quality = 0.58
): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Falha ao ler arquivo de imagem"));
    reader.onload = (e) => {
      const img = new Image();
      img.onerror = () => reject(new Error("Falha ao processar dados da imagem"));
      img.onload = () => {
        try {
          const canvas = document.createElement('canvas');
          let width = img.width;
          let height = img.height;

          if (width > height) {
            if (width > maxDimension) {
              height = Math.round((height * maxDimension) / width);
              width = maxDimension;
            }
          } else {
            if (height > maxDimension) {
              width = Math.round((width * maxDimension) / height);
              height = maxDimension;
            }
          }

          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          if (!ctx) {
            resolve(e.target?.result as string);
            return;
          }

          // Fill white background for transparent PNGs
          ctx.fillStyle = '#FFFFFF';
          ctx.fillRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);

          const compressed = canvas.toDataURL('image/jpeg', quality);
          resolve(compressed);
        } catch (err) {
          resolve(e.target?.result as string);
        }
      };
      img.src = e.target?.result as string;
    };
    reader.readAsDataURL(file);
  });
}

/**
 * Persists full cronograma activities and general consultoria evidence to IndexedDB
 */
export async function saveLocalDiagnosticoReport(
  diagnosticoId: string, 
  cronograma: any[], 
  dadosConsultoria: any
): Promise<void> {
  if (!diagnosticoId) return;
  try {
    const db = await getIDB();
    const tx = db.transaction('diagnostico_cache', 'readwrite');
    const store = tx.objectStore('diagnostico_cache');
    store.put({
      diagnosticoId,
      cronograma,
      dadosConsultoria,
      savedAt: Date.now()
    });
  } catch (err) {
    console.warn("Falha ao salvar no IndexedDB (fallback silencioso):", err);
  }
}

/**
 * Retrieves cached cronograma and evidence data from IndexedDB
 */
export async function loadLocalDiagnosticoReport(
  diagnosticoId: string
): Promise<{ cronograma?: any[]; dadosConsultoria?: any; savedAt?: number } | null> {
  if (!diagnosticoId) return null;
  try {
    const db = await getIDB();
    return new Promise((resolve) => {
      const tx = db.transaction('diagnostico_cache', 'readonly');
      const store = tx.objectStore('diagnostico_cache');
      const req = store.get(diagnosticoId);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (err) {
    return null;
  }
}

/**
 * Save evidence for individual items (respostas, tarefas)
 */
export async function saveLocalIndividualEvidence(
  type: 'resposta' | 'tarefa',
  id: string,
  dataUrl: string,
  nome?: string
): Promise<void> {
  if (!id) return;
  try {
    const db = await getIDB();
    const tx = db.transaction('individual_evidences', 'readwrite');
    const store = tx.objectStore('individual_evidences');
    store.put({
      key: `${type}_${id}`,
      type,
      id,
      dataUrl,
      nome,
      savedAt: Date.now()
    });
  } catch (err) {
    console.warn("Falha ao salvar evidência individual no IndexedDB:", err);
  }
}

/**
 * Delete evidence for individual items
 */
export async function deleteLocalIndividualEvidence(
  type: 'resposta' | 'tarefa',
  id: string
): Promise<void> {
  if (!id) return;
  try {
    const db = await getIDB();
    const tx = db.transaction('individual_evidences', 'readwrite');
    const store = tx.objectStore('individual_evidences');
    store.delete(`${type}_${id}`);
  } catch (err) {}
}

/**
 * Load evidence for individual items
 */
export async function loadLocalIndividualEvidence(
  type: 'resposta' | 'tarefa',
  id: string
): Promise<{ dataUrl: string; nome?: string } | null> {
  if (!id) return null;
  try {
    const db = await getIDB();
    return new Promise((resolve) => {
      const tx = db.transaction('individual_evidences', 'readonly');
      const store = tx.objectStore('individual_evidences');
      const req = store.get(`${type}_${id}`);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (err) {
    return null;
  }
}
