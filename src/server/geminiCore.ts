import { GoogleGenAI } from '@google/genai';
import {
  GEMINI_MODEL_CHAIN,
  GEMINI_PRIMARY_MODEL,
  cleanApiKey,
  friendlyGeminiError,
  isValidGeminiApiKey,
  normalizeGeminiModel,
} from '../lib/ai/geminiModels';

/**
 * Lógica do proxy do Gemini, usada tanto pela Netlify Function (netlify/functions/gemini.mts)
 * quanto pelo server.ts (Docker / desenvolvimento local).
 *
 * A chave do sistema vem SOMENTE da variável de ambiente GEMINI_API_KEY — nunca do código.
 *   - Netlify: Site configuration → Environment variables → GEMINI_API_KEY (escopo: Functions)
 *   - Local/Docker: arquivo .env (que está no .gitignore)
 * Opcional: GEMINI_MODEL para trocar o modelo principal sem mexer no código.
 */

export interface CoreResult {
  status: number;
  body: any;
}

export const isValidApiKey = isValidGeminiApiKey;

export function getSystemGeminiApiKey(): string {
  const direct = cleanApiKey(process.env.GEMINI_API_KEY);
  if (isValidGeminiApiKey(direct)) return direct;
  // Aceita variações de nome já usadas no painel (ex.: Gemini_api_kei, GOOGLE_API_KEY).
  for (const envKey of Object.keys(process.env)) {
    const l = envKey.toLowerCase();
    if (l.startsWith('vite_')) continue; // variáveis VITE_* vão para o navegador: nunca usar para segredo
    if (l.includes('gemini_api') || l.includes('gemini_key') || l.includes('gemini_kei') || l === 'google_api_key') {
      const val = cleanApiKey(process.env[envKey]);
      if (isValidGeminiApiKey(val)) return val;
    }
  }
  return '';
}

function resolveKey(customKey: unknown): { key: string; source: 'user' | 'system' | 'none' } {
  const custom = cleanApiKey(customKey);
  if (isValidGeminiApiKey(custom)) return { key: custom, source: 'user' };
  const sys = getSystemGeminiApiKey();
  return sys ? { key: sys, source: 'system' } : { key: '', source: 'none' };
}

function primaryModel(): string {
  return normalizeGeminiModel(process.env.GEMINI_MODEL, GEMINI_PRIMARY_MODEL);
}

function modelChain(requested?: unknown): string[] {
  const primary = primaryModel();
  return Array.from(new Set([normalizeGeminiModel(requested, primary), primary, ...GEMINI_MODEL_CHAIN]));
}

function withTimeout<T>(p: Promise<T>, ms: number, msg: string): Promise<T> {
  let t: ReturnType<typeof setTimeout>;
  return Promise.race([
    p.finally(() => clearTimeout(t)),
    new Promise<T>((_, reject) => { t = setTimeout(() => reject(new Error(msg)), ms); }),
  ]);
}

/** Erros de chave/cota não melhoram trocando de modelo: para de tentar na hora. */
function isAuthOrQuotaError(e: any): boolean {
  const s = e?.status ?? e?.code;
  const m = String(e?.message || '').toLowerCase();
  return s === 401 || s === 403 || s === 429 || m.includes('api key') || m.includes('unauthenticated') || m.includes('permission_denied') || m.includes('resource_exhausted');
}

function errorStatus(e: any): number {
  const s = Number(e?.status ?? e?.code);
  return s === 401 || s === 403 || s === 429 ? s : 500;
}

export function geminiConfig(): CoreResult {
  return { status: 200, body: { hasSystemKey: !!getSystemGeminiApiKey(), model: primaryModel() } };
}

export async function geminiTest(customKey: unknown): Promise<CoreResult> {
  const { key, source } = resolveKey(customKey);
  if (!key) return { status: 400, body: { ok: false, error: 'Chave API do Gemini não configurada no servidor (variável GEMINI_API_KEY).' } };
  const ai = new GoogleGenAI({ apiKey: key });
  let lastErr: any = null;
  for (const model of modelChain()) {
    try {
      const r: any = await withTimeout(ai.models.generateContent({ model, contents: "Responda apenas 'OK'." }), 10000, 'Timeout de teste da API');
      if (r?.text) {
        const origem = source === 'user' ? 'sua chave' : 'a chave do sistema';
        return { status: 200, body: { ok: true, model, source, message: `Conexão bem-sucedida usando ${origem} com o modelo ${model}!` } };
      }
    } catch (e) {
      lastErr = e;
      if (isAuthOrQuotaError(e)) break;
    }
  }
  return { status: errorStatus(lastErr), body: { ok: false, error: friendlyGeminiError(lastErr) } };
}

export async function geminiGenerate(body: any, customKey: unknown): Promise<CoreResult> {
  const { model, contents, config } = body || {};
  if (contents === undefined || contents === null || contents === '') {
    return { status: 400, body: { error: 'Conteúdo vazio.' } };
  }
  const { key } = resolveKey(customKey);
  if (!key) {
    return {
      status: 400,
      body: { error: 'Chave API do Gemini não configurada no servidor. Cadastre GEMINI_API_KEY nas variáveis do Netlify ou use sua própria chave em Configurações.' },
    };
  }

  const ai = new GoogleGenAI({ apiKey: key });
  let lastErr: any = null;
  // Netlify Functions síncronas têm limite de ~26s no total: reparte esse tempo entre as tentativas.
  const deadline = Date.now() + 23000;
  for (const m of modelChain(model)) {
    const remaining = deadline - Date.now();
    if (remaining < 2000) break;
    try {
      const response: any = await withTimeout(
        ai.models.generateContent({ model: m, contents, config }),
        Math.min(20000, remaining),
        'Tempo limite de resposta do Gemini excedido'
      );
      if (response && typeof response.text === 'string' && response.text !== '') {
        return { status: 200, body: { text: response.text, model: m } };
      }
      lastErr = new Error(`O modelo ${m} retornou resposta vazia.`);
    } catch (err: any) {
      console.warn(`[Gemini Proxy] Modelo ${m} falhou:`, err?.message || err);
      lastErr = err;
      if (isAuthOrQuotaError(err)) break;
    }
  }
  return { status: errorStatus(lastErr), body: { error: friendlyGeminiError(lastErr) } };
}
