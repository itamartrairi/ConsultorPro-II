import { GoogleGenAI } from '@google/genai';

/**
 * Lógica do proxy do Gemini, usada tanto pela Netlify Function (netlify/functions/gemini.mts)
 * quanto pelo server.ts (Docker / desenvolvimento local).
 *
 * A chave do sistema vem SOMENTE da variável de ambiente GEMINI_API_KEY.
 */

export interface CoreResult {
  status: number;
  body: any;
}

const GEMINI_CHAIN = ['gemini-2.5-flash', 'gemini-2.0-flash', 'gemini-1.5-flash'];
const GEMINI_PRIMARY = GEMINI_CHAIN[0];

export function isValidApiKey(val: unknown): val is string {
  if (typeof val !== 'string') return false;
  const t = val.trim();
  return t !== '' && t !== 'undefined' && t !== 'null' && (t.startsWith('AIzaSy') || t.startsWith('AQ.') || t.length >= 15);
}

export function getSystemGeminiApiKey(): string {
  const direct = process.env.GEMINI_API_KEY;
  if (isValidApiKey(direct)) return direct.trim();
  // Aceita variações de nome já usadas no painel (ex.: Gemini_api_kei).
  for (const envKey of Object.keys(process.env)) {
    const l = envKey.toLowerCase();
    if (l.includes('gemini_api') || l.includes('gemini_key') || l.includes('gemini_kei')) {
      const val = process.env[envKey];
      if (isValidApiKey(val)) return val.trim();
    }
  }
  return '';
}

function resolveKey(customKey: unknown): string {
  return isValidApiKey(customKey) ? customKey.trim() : getSystemGeminiApiKey();
}

function normalizeModel(m?: unknown): string {
  if (typeof m !== 'string' || !m) return GEMINI_PRIMARY;
  // Só modelos Gemini: impede usar o proxy para outros fins.
  return m.startsWith('gemini-') ? m : GEMINI_PRIMARY;
}

function withTimeout<T>(p: Promise<T>, ms: number, msg: string): Promise<T> {
  return Promise.race([p, new Promise<T>((_, reject) => setTimeout(() => reject(new Error(msg)), ms))]);
}

export function geminiConfig(): CoreResult {
  return { status: 200, body: { hasSystemKey: !!getSystemGeminiApiKey() } };
}

export async function geminiTest(customKey: unknown): Promise<CoreResult> {
  const apiKey = resolveKey(customKey);
  if (!apiKey) return { status: 400, body: { ok: false, error: 'Chave API do Gemini não configurada no servidor.' } };
  const ai = new GoogleGenAI({ apiKey });
  let lastErr: any = null;
  for (const model of GEMINI_CHAIN) {
    try {
      const r: any = await withTimeout(ai.models.generateContent({ model, contents: "Responda apenas 'OK'." }), 6000, 'Timeout de teste da API');
      if (r?.text) return { status: 200, body: { ok: true, model, message: `Conexão bem-sucedida com o modelo ${model}!` } };
    } catch (e) {
      lastErr = e;
    }
  }
  return { status: 500, body: { ok: false, error: lastErr?.message || 'Falha ao conectar aos modelos do Gemini.' } };
}

export async function geminiGenerate(body: any, customKey: unknown): Promise<CoreResult> {
  const { model, contents, config } = body || {};
  if (contents === undefined || contents === null || contents === '') {
    return { status: 400, body: { error: 'Conteúdo vazio.' } };
  }
  const apiKey = resolveKey(customKey);
  if (!apiKey) {
    return {
      status: 400,
      body: { error: 'Chave API do Gemini não configurada no servidor. Cadastre GEMINI_API_KEY nas variáveis do Netlify ou use sua própria chave em Configurações.' },
    };
  }

  const ai = new GoogleGenAI({ apiKey });
  const candidates = Array.from(new Set([normalizeModel(model), ...GEMINI_CHAIN]));
  let lastErr: any = null;
  // Netlify Functions síncronas têm limite de ~26s no total: reparte esse tempo entre as tentativas.
  const deadline = Date.now() + 22000;
  for (const m of candidates) {
    const remaining = deadline - Date.now();
    if (remaining < 2000) break;
    try {
      const response: any = await withTimeout(
        ai.models.generateContent({ model: m, contents, config }),
        Math.min(15000, remaining),
        'Tempo limite de resposta do Gemini excedido'
      );
      if (response && response.text !== undefined) {
        return { status: 200, body: { text: response.text } };
      }
    } catch (err: any) {
      console.warn(`[Gemini Proxy] Modelo ${m} falhou:`, err?.message || err);
      lastErr = err;
    }
  }
  return {
    status: 500,
    body: { error: lastErr?.message || 'Não foi possível conectar aos modelos do Gemini no momento. Tente novamente mais tarde.' },
  };
}
