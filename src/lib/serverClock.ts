import { apiUrl } from './gemini';
import { setServerTime } from './licenseDays';

/**
 * Busca a hora do servidor (Netlify Function /api/time) e passa para o cálculo da
 * licença. Se falhar (sem internet ou no ambiente de desenvolvimento), o cálculo segue
 * com o relógio local, mas sem nunca voltar no tempo (ver trustedNow em licenseDays.ts).
 */
export async function syncServerClock(timeoutMs = 8000): Promise<boolean> {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeoutMs);
  try {
    const started = Date.now();
    const res = await fetch(apiUrl('/api/time'), { cache: 'no-store', signal: ctrl.signal });
    if (!res.ok) return false;
    const data = await res.json();
    const serverMs = Number(data?.now);
    if (!Number.isFinite(serverMs) || serverMs <= 0) return false;
    // Compensa metade do tempo de ida e volta da requisição.
    setServerTime(serverMs + (Date.now() - started) / 2);
    return true;
  } catch {
    return false;
  } finally {
    clearTimeout(timer);
  }
}
