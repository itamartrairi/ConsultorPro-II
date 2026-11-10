/**
 * Sanitiza objetos antes de gravar no Firestore (remove undefined, preserva Timestamps).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

export function sanitizeForFirestore(obj: any): any {
  if (obj === null || typeof obj !== 'object') return obj;
  if (obj instanceof Date) return obj.toISOString();
  if (
    // Marcadores do Firestore (serverTimestamp, deleteField, increment...). No build de
    // produção o nome da classe é minificado, então o teste por constructor.name falhava e o
    // marcador virava um objeto comum ({ _methodName }): a data de cadastro não era gravada.
    typeof obj._methodName === 'string' ||
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
