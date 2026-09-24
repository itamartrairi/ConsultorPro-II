/**
 * Modelos do Gemini usados pelo app (compartilhado entre servidor e navegador).
 *
 * - gemini-3.6-flash      → modelo principal (GA/estável).
 * - gemini-3.5-flash-lite → reserva rápida e barata, usada se o principal falhar.
 *
 * Modelos antigos (1.5, 2.0, 2.5) já foram desligados pelo Google e respondem 404;
 * por isso qualquer pedido para eles é redirecionado ao modelo principal.
 */
export const GEMINI_PRIMARY_MODEL = 'gemini-3.6-flash';
export const GEMINI_FALLBACK_MODEL = 'gemini-3.5-flash-lite';
export const GEMINI_MODEL_CHAIN = [GEMINI_PRIMARY_MODEL, GEMINI_FALLBACK_MODEL];

const RETIRED_MODEL = /^gemini-(1\.|2\.0|2\.5)/;

export function normalizeGeminiModel(m?: unknown, primary: string = GEMINI_PRIMARY_MODEL): string {
  if (typeof m !== 'string') return primary;
  const t = m.trim();
  // Só modelos Gemini: impede usar o proxy para outros fins.
  if (!t.startsWith('gemini-')) return primary;
  if (RETIRED_MODEL.test(t)) return primary;
  return t;
}

/** Tira espaços, aspas e quebras de linha que costumam vir junto ao colar a chave. */
export function cleanApiKey(val: unknown): string {
  if (typeof val !== 'string') return '';
  return val.trim().replace(/^["'`]+|["'`]+$/g, '').replace(/\s+/g, '');
}

/**
 * Aceita os dois formatos de chave do Google AI Studio:
 *  - "AQ.…"   (formato novo, padrão desde 2026)
 *  - "AIza…"  (formato antigo)
 */
export function isValidGeminiApiKey(val: unknown): boolean {
  const t = cleanApiKey(val);
  if (!t || t === 'undefined' || t === 'null') return false;
  return t.startsWith('AQ.') || t.startsWith('AIza') || t.length >= 30;
}

/** Traduz os erros mais comuns da API do Google para mensagens claras. */
export function friendlyGeminiError(raw: unknown, status?: number): string {
  const msg = String((raw as any)?.message ?? raw ?? '');
  const low = msg.toLowerCase();
  const code = status ?? (raw as any)?.status;
  if (code === 401 || code === 403 || low.includes('api key not valid') || low.includes('api_key_invalid') || low.includes('unauthenticated') || low.includes('permission_denied')) {
    return 'A chave do Gemini foi recusada pelo Google (inválida, revogada ou sem permissão para a Gemini API). Gere uma nova chave no Google AI Studio.';
  }
  if (code === 429 || low.includes('resource_exhausted') || low.includes('quota')) {
    return 'Limite de uso da chave do Gemini atingido. Aguarde alguns minutos ou ative o faturamento no projeto do Google AI Studio.';
  }
  if (low.includes('timeout') || low.includes('tempo limite')) {
    return 'O Gemini demorou demais para responder. Tente novamente.';
  }
  return msg || 'Falha ao conectar ao Gemini.';
}
