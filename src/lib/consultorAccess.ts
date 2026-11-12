/**
 * Acesso dos consultores de uma credenciada.
 *
 * Cada consultor cadastrado com e-mail Google ganha um documento em
 * `consultores_acesso/{email em minúsculas}` apontando para o dono da credenciada.
 * No login, o app procura esse documento: se existir e estiver ativo, o consultor
 * trabalha nos dados da credenciada (mesmo ownerId) e usa a licença dela.
 *
 * As regras do Firestore (firestore.rules) usam o mesmo documento para liberar o acesso.
 */
import { doc, getDoc, setDoc, deleteDoc, serverTimestamp } from 'firebase/firestore';
import { db } from '../firebase';
import type { ConsultorCredenciada, EmpresaCredenciada } from '../types/domain';

export const CONSULTORES_ACESSO = 'consultores_acesso';

export interface ConsultorAcesso {
  email: string;
  ownerId: string;
  credenciadaId: string;
  credenciadaNome: string;
  consultorId: string;
  nome: string;
  ativo: boolean;
}

export function normalizeEmail(email?: string | null): string {
  return (email || '').trim().toLowerCase();
}

export function isValidEmail(email?: string | null): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail(email));
}

export function novoConsultorId(): string {
  return 'cons_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/** Acesso de consultor do e-mail logado (null se não houver ou estiver inativo). */
export async function fetchConsultorAcesso(email?: string | null): Promise<ConsultorAcesso | null> {
  const key = normalizeEmail(email);
  if (!key) return null;
  const snap = await getDoc(doc(db, CONSULTORES_ACESSO, key));
  if (!snap.exists()) return null;
  const data = snap.data() as ConsultorAcesso;
  if (!data.ativo || !data.ownerId) return null;
  return { ...data, email: key };
}

/**
 * Grava/remove os acessos depois de salvar os consultores de uma credenciada.
 * Só o dono da credenciada (ou o admin) consegue gravar: as regras conferem o ownerId.
 * Retorna os e-mails que não puderam ser liberados (ex.: já são consultores de outra credenciada).
 */
export async function syncConsultoresAcesso(
  cred: Pick<EmpresaCredenciada, 'id' | 'razaoSocial' | 'ownerId'>,
  consultores: ConsultorCredenciada[] = [],
  anteriores: ConsultorCredenciada[] = []
): Promise<string[]> {
  const falhas: string[] = [];
  const atuais = new Map<string, ConsultorCredenciada>();
  consultores.forEach(c => {
    const e = normalizeEmail(c.email);
    if (e && isValidEmail(e)) atuais.set(e, c);
  });

  for (const [email, c] of atuais) {
    try {
      await setDoc(doc(db, CONSULTORES_ACESSO, email), {
        email,
        ownerId: cred.ownerId,
        credenciadaId: cred.id,
        credenciadaNome: cred.razaoSocial || '',
        consultorId: c.id,
        nome: c.nome || '',
        ativo: c.ativo !== false,
        updatedAt: serverTimestamp(),
      });
    } catch (e) {
      console.error('Não foi possível liberar o acesso do consultor', email, e);
      falhas.push(email);
    }
  }

  const removidos = anteriores
    .map(c => normalizeEmail(c.email))
    .filter(e => e && !atuais.has(e));
  for (const email of Array.from(new Set(removidos))) {
    try {
      await deleteDoc(doc(db, CONSULTORES_ACESSO, email));
    } catch (e) {
      console.error('Não foi possível remover o acesso do consultor', email, e);
    }
  }
  return falhas;
}

/** Opções de consultor de uma credenciada: o principal (cadastro) e os adicionais ativos. */
export interface OpcaoConsultor {
  id: string;
  nome: string;
  email?: string;
  cpf?: string;
}

export const CONSULTOR_PRINCIPAL_ID = 'principal';

export function consultoresDaCredenciada(cred?: EmpresaCredenciada | null): OpcaoConsultor[] {
  if (!cred) return [];
  const lista: OpcaoConsultor[] = [];
  if ((cred.consultor || '').trim()) {
    lista.push({ id: CONSULTOR_PRINCIPAL_ID, nome: cred.consultor!.trim(), email: cred.email, cpf: cred.cpfConsultor });
  }
  (cred.consultores || [])
    .filter(c => c && c.ativo !== false && (c.nome || '').trim())
    .forEach(c => lista.push({ id: c.id, nome: c.nome.trim(), email: c.email, cpf: c.cpf }));
  return lista;
}

/** Consultor padrão para um atendimento: o próprio consultor logado, senão o principal. */
export function consultorPadrao(cred: EmpresaCredenciada | null | undefined, loginEmail?: string | null): OpcaoConsultor | null {
  const opcoes = consultoresDaCredenciada(cred);
  const e = normalizeEmail(loginEmail);
  const proprio = e ? opcoes.find(o => o.id !== CONSULTOR_PRINCIPAL_ID && normalizeEmail(o.email) === e) : undefined;
  return proprio || opcoes[0] || null;
}
