/**
 * Camada fina sobre o 'firebase/firestore' que garante o isolamento de dados por usuário.
 *
 * As regras do Firestore (firestore.rules) só permitem ler/gravar documentos cujo
 * `ownerId` é o uid do usuário logado. Este módulo:
 *   - completa `ownerId` em toda gravação (addDoc / setDoc / batch.set) que não o tenha
 *     ou que ainda esteja como 'local' / 'local_user' (dados criados offline);
 *   - oferece `ownerFilter()` para acrescentar `where('ownerId', '==', uid)` às consultas
 *     (o Firestore recusa consultas que possam retornar documentos de outros usuários).
 *
 * Use sempre `import { ... } from './lib/firestoreOwned'` no lugar de 'firebase/firestore'.
 */
import * as fs from 'firebase/firestore';
import { getAuth } from 'firebase/auth';

export * from 'firebase/firestore';

// Coleções que pertencem a um usuário. 'system_health' é só o teste de conexão.
const OWNED_COLLECTIONS = new Set([
  'empresas',
  'diagnosticos',
  'respostas',
  'tarefas_plano',
  'empresas_credenciadas',
  'agenda_eventos',
  'disc_avaliacoes',
  'maturidade_avaliacoes',
  'analises_resultado',
  'premissas',
  'problemas',
  'solucoes',
  'areas',
  'segmentos',
]);

const PLACEHOLDER_OWNERS = new Set(['', 'local', 'local_user']);

function currentUid(): string | null {
  try {
    return getAuth().currentUser?.uid ?? null;
  } catch {
    return null;
  }
}

function rootCollectionOf(path: string): string {
  return path.split('/')[0];
}

function withOwner<T>(collectionPath: string, data: T): T {
  if (!data || typeof data !== 'object' || Array.isArray(data)) return data;
  if (!OWNED_COLLECTIONS.has(rootCollectionOf(collectionPath))) return data;
  const uid = currentUid();
  if (!uid) return data;
  const current = (data as any).ownerId;
  if (typeof current === 'string' && !PLACEHOLDER_OWNERS.has(current)) return data;
  return { ...(data as any), ownerId: uid };
}

/** Filtro obrigatório nas consultas às coleções isoladas por usuário. */
export function ownerFilter(): fs.QueryFieldFilterConstraint {
  return fs.where('ownerId', '==', currentUid() ?? '__sem_login__');
}

export function addDoc<A, D extends fs.DocumentData>(
  reference: fs.CollectionReference<A, D>,
  data: fs.WithFieldValue<A>
): Promise<fs.DocumentReference<A, D>> {
  return fs.addDoc(reference, withOwner(reference.path, data));
}

export function setDoc<A, D extends fs.DocumentData>(
  reference: fs.DocumentReference<A, D>,
  data: any,
  options?: fs.SetOptions
): Promise<void> {
  const payload = withOwner(reference.parent.path, data);
  return options ? fs.setDoc(reference, payload, options) : fs.setDoc(reference, payload);
}

export function writeBatch(firestore: fs.Firestore): fs.WriteBatch {
  const batch = fs.writeBatch(firestore);
  const originalSet = batch.set.bind(batch) as any;
  (batch as any).set = (reference: fs.DocumentReference<any, any>, data: any, options?: fs.SetOptions) => {
    const payload = withOwner(reference.parent.path, data);
    return options ? originalSet(reference, payload, options) : originalSet(reference, payload);
  };
  return batch;
}
