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

/**
 * Dias restantes da licença de uma credenciada. Nunca retorna NaN.
 * Definitiva → 99999. Com validadeLicenca → dias até a validade.
 * Senão → limite (diasTeste ou padrão do plano) menos os dias desde o cadastro.
 */
export function computeLicenseDaysLeft(emp: any): number {
  const plano = emp?.tipoPlano || 'Teste';
  if (plano === 'Definitiva') return 99999;
  const validade = toJsDate(emp?.validadeLicenca);
  if (validade) return Math.ceil((validade.getTime() - Date.now()) / DAY_MS);
  const diasTeste = Number(emp?.diasTeste);
  const limitDays = Number.isFinite(diasTeste) && diasTeste > 0 ? diasTeste : (plano === 'Mensal' ? 30 : plano === 'Anual' ? 365 : 30);
  const cadastro = toJsDate(emp?.dataCadastro);
  if (!cadastro) return limitDays;
  return limitDays - Math.floor((Date.now() - cadastro.getTime()) / DAY_MS);
}
