/**
 * Converte qualquer formato de data usado no app em Date (ou null se não der):
 * Timestamp do Firestore, {seconds, nanoseconds} (cópia salva no navegador),
 * {_seconds} (Firebase Admin), número (ms ou s), texto ISO ou "dd/mm/aaaa".
 */
export function toJsDate(v: any): Date | null {
  if (v === null || v === undefined || v === '') return null;
  try {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    if (typeof v === 'object') {
      if (typeof v.toDate === 'function') { const d = v.toDate(); return d instanceof Date && !isNaN(d.getTime()) ? d : null; }
      const secs = typeof v.seconds === 'number' ? v.seconds : typeof v._seconds === 'number' ? v._seconds : null;
      if (secs !== null) return new Date(secs * 1000);
      return null;
    }
    if (typeof v === 'number') {
      if (!isFinite(v) || v <= 0) return null;
      return new Date(v < 10000000000 ? v * 1000 : v);
    }
    if (typeof v === 'string') {
      const t = v.trim();
      const br = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (br) {
        const d = new Date(Number(br[3]), Number(br[2]) - 1, Number(br[1]), 12);
        return isNaN(d.getTime()) ? null : d;
      }
      const d = new Date(/^\d{4}-\d{2}-\d{2}$/.test(t) ? t + 'T12:00:00' : t);
      return isNaN(d.getTime()) ? null : d;
    }
  } catch {}
  return null;
}

const DAY_MS = 1000 * 60 * 60 * 24;

// --- Relógio da licença ------------------------------------------------------
// Os dias da licença não podem depender só do relógio do computador do usuário:
// atrasar o relógio do PC dava dias a mais. Usamos a hora do servidor quando ela é
// conhecida (setServerTime) e, sem ela, nunca deixamos o tempo voltar para antes do
// último horário já visto neste navegador.
const CLOCK_KEY = 'cp_license_clock';
let serverOffsetMs: number | null = null;

function readLastSeen(): number {
  try {
    const v = Number((globalThis as any).localStorage?.getItem(CLOCK_KEY));
    return Number.isFinite(v) && v > 0 ? v : 0;
  } catch {
    return 0;
  }
}

function saveLastSeen(ms: number): void {
  try {
    (globalThis as any).localStorage?.setItem(CLOCK_KEY, String(ms));
  } catch {
    /* sem armazenamento local (ex.: testes no Node) */
  }
}

/** Registra a hora do servidor (ms). Depois disso o relógio do PC só conta a diferença. */
export function setServerTime(serverMs: number): void {
  if (!Number.isFinite(serverMs) || serverMs <= 0) return;
  serverOffsetMs = serverMs - Date.now();
  saveLastSeen(serverMs);
}

/** Hora usada nos cálculos de licença (servidor quando conhecida; nunca volta no tempo). */
export function trustedNow(): number {
  const now = serverOffsetMs !== null ? Date.now() + serverOffsetMs : Math.max(Date.now(), readLastSeen());
  saveLastSeen(now);
  return now;
}

/** Duração de cada plano pago, em dias. */
export function licensePeriodDays(plano: string): number {
  return plano === 'Anual' ? 365 : 30;
}

export function isPaidPlan(plano: unknown): boolean {
  return plano === 'Mensal' || plano === 'Anual' || plano === 'Definitiva';
}

/** Data (aaaa-mm-dd) daqui a N dias a partir de uma data base (padrão: hoje). */
export function addDaysIso(days: number, base: Date = new Date()): string {
  const d = new Date(base.getTime() + days * DAY_MS);
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/**
 * Dias restantes da licença de uma credenciada. Nunca retorna NaN.
 * Definitiva → 99999. Com validadeLicenca → dias até a validade.
 * Senão → limite menos os dias desde o cadastro. O limite vem de `diasTeste` SÓ no
 * plano Teste; nos planos pagos vale a duração do plano (30 ou 365 dias). Antes, o
 * `diasTeste: 30` criado no cadastro fazia o plano Anual vencer em 30 dias.
 */
export function computeLicenseDaysLeft(emp: any): number {
  const plano = emp?.tipoPlano || 'Teste';
  if (plano === 'Definitiva') return 99999;
  const validade = toJsDate(emp?.validadeLicenca);
  if (validade) return Math.ceil((validade.getTime() - trustedNow()) / DAY_MS);
  const diasTeste = Number(emp?.diasTeste);
  const limitDays = plano === 'Teste' && Number.isFinite(diasTeste) && diasTeste > 0 ? diasTeste : licensePeriodDays(plano);
  const cadastro = toJsDate(emp?.dataCadastro);
  if (!cadastro) return limitDays;
  return limitDays - Math.floor((trustedNow() - cadastro.getTime()) / DAY_MS);
}
