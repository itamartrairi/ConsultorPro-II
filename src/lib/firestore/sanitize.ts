/**
 * Sanitiza objetos antes de gravar no Firestore (remove undefined, preserva Timestamps).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export function sanitizeForFirestore(obj: any): any {
  if (obj === null || typeof obj !== 'object') return obj;
  if (obj instanceof Date) return obj.toISOString();
  if (
    typeof obj.toDate === 'function' ||
    typeof obj.toMillis === 'function' ||
    obj.constructor?.name === 'Timestamp' ||
    obj.constructor?.name === 'FieldValue' ||
    obj._delegate
  ) {
    return obj;
  }
  if (Array.isArray(obj)) return obj.map((v) => sanitizeForFirestore(v));

  const result: any = {};
  for (const key in obj) {
    if (obj[key] !== undefined) {
      result[key] = sanitizeForFirestore(obj[key]);
    }
  }
  return result;
}
