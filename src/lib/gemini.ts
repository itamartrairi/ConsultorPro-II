import { auth } from '../firebase';
import {
  GEMINI_MODEL_CHAIN,
  GEMINI_PRIMARY_MODEL,
  cleanApiKey,
  friendlyGeminiError,
  isValidGeminiApiKey as isValidKey,
  normalizeGeminiModel,
} from './ai/geminiModels';

/**
 * Cliente do Gemini no navegador.
 *
 * A chave do sistema (GEMINI_API_KEY) NUNCA chega ao navegador: ela fica só no servidor
 * (Netlify Function /api/gemini/*). O navegador só conhece a chave OPCIONAL que o próprio
 * usuário salvar em Configurações (localStorage 'custom_gemini_api_key').
 */

export function isValidGeminiApiKey(k: any): boolean {
  return isValidKey(k);
}

// Prefixos de chaves que ficaram públicas — removidas do navegador de quem as salvou.
const LEAKED_GEMINI_KEY_PREFIXES = ['AQ.Ab8RN6LGK4um2g'];

export function readCustomGeminiKey(): string {
  try {
    const raw = cleanApiKey(localStorage.getItem('custom_gemini_api_key') || '');
    if (LEAKED_GEMINI_KEY_PREFIXES.some((p) => raw.startsWith(p))) {
      localStorage.removeItem('custom_gemini_api_key');
      return '';
    }
    return isValidKey(raw) ? raw : '';
  } catch {
    return '';
  }
}

// Token de login do Firebase, exigido pelo proxy do Gemini (evita uso da chave por terceiros).
export async function geminiAuthHeaders(): Promise<Record<string, string>> {
  try {
    const token = await auth.currentUser?.getIdToken();
    return token ? { Authorization: `Bearer ${token}` } : {};
  } catch {
    return {};
  }
}

// No app desktop (Electron) a página abre como arquivo local e não existe "/api".
// Defina VITE_API_BASE_URL=https://SEU-SITE.netlify.app no build do desktop para usar
// as Netlify Functions do site publicado.
export function apiUrl(path: string): string {
  const base = ((import.meta as any).env?.VITE_API_BASE_URL || '').replace(/\/+$/, '');
  if (base && typeof window !== 'undefined' && window.location.protocol === 'file:') {
    return base + path;
  }
  return path;
}

function toRestContents(contents: any): any[] {
  if (typeof contents === 'string') return [{ role: 'user', parts: [{ text: contents }] }];
  if (Array.isArray(contents)) {
    // Lista de "contents" ({role, parts}) ou lista de "parts" ({text}/{inlineData}).
    return contents.length && contents[0]?.parts ? contents : [{ role: 'user', parts: contents }];
  }
  if (contents?.parts) return [{ role: contents.role || 'user', ...contents }];
  return [{ role: 'user', parts: [{ text: JSON.stringify(contents) }] }];
}

/**
 * Chamada DIRETA ao Google, usada só como contingência quando o servidor está fora do ar
 * e o usuário salvou a própria chave. A chave vai no cabeçalho x-goog-api-key (formato
 * exigido pelas chaves novas "AQ.") e nunca na URL.
 */
export async function directGeminiGenerate(
  apiKey: string,
  model: string | undefined,
  contents: any,
  config?: any
): Promise<{ text: string; model: string }> {
  const key = cleanApiKey(apiKey);
  const chain = Array.from(new Set([normalizeGeminiModel(model), ...GEMINI_MODEL_CHAIN]));
  let lastErr: any = null;

  for (const m of chain) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 25000);
    try {
      const { responseSchema, responseMimeType, systemInstruction, ...rest } = config || {};
      const body: any = { contents: toRestContents(contents) };
      const generationConfig: any = { ...rest };
      if (responseMimeType) generationConfig.responseMimeType = responseMimeType;
      if (responseSchema) generationConfig.responseSchema = responseSchema;
      if (Object.keys(generationConfig).length) body.generationConfig = generationConfig;
      if (systemInstruction) {
        body.systemInstruction = typeof systemInstruction === 'string' ? { parts: [{ text: systemInstruction }] } : systemInstruction;
      }

      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-goog-api-key': key },
        signal: controller.signal,
        body: JSON.stringify(body),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        const err: any = new Error(data?.error?.message || `Erro HTTP ${res.status}`);
        err.status = res.status;
        throw err;
      }
      const text = (data?.candidates?.[0]?.content?.parts || [])
        .map((p: any) => (p?.thought ? '' : p?.text || ''))
        .join('');
      if (text) return { text, model: m };
      lastErr = new Error(`O modelo ${m} retornou resposta vazia.`);
    } catch (e: any) {
      lastErr = e?.name === 'AbortError' ? new Error('Tempo limite de resposta do Gemini excedido') : e;
      const s = lastErr?.status;
      if (s === 400 || s === 401 || s === 403 || s === 429) break; // trocar de modelo não resolve
    } finally {
      clearTimeout(timeoutId);
    }
  }
  throw new Error(friendlyGeminiError(lastErr, lastErr?.status));
}

/**
 * Gera texto com o Gemini. Sempre tenta primeiro o servidor (que usa a chave GEMINI_API_KEY
 * guardada nas variáveis de ambiente, ou a chave própria do usuário, se ele salvou uma).
 */
export async function callGemini({ model, contents, config }: { model?: string; contents: any; config?: any }): Promise<{ text: string }> {
  const customKey = readCustomGeminiKey();
  const headers: Record<string, string> = { 'Content-Type': 'application/json', ...(await geminiAuthHeaders()) };
  if (customKey) headers['x-custom-api-key'] = customKey;

  let proxyError: Error | null = null;
  try {
    const res = await fetch(apiUrl('/api/gemini/generate'), {
      method: 'POST',
      headers,
      body: JSON.stringify({ model: normalizeGeminiModel(model), contents, config }),
    });
    const isJson = (res.headers.get('content-type') || '').includes('application/json');
    const data: any = isJson ? await res.json().catch(() => ({})) : {};
    if (res.ok && typeof data?.text === 'string') return { text: data.text };
    // Resposta HTML/404 = não há servidor (ex.: app desktop sem VITE_API_BASE_URL).
    const serverMissing = !isJson || res.status === 404 || res.status === 502 || res.status === 503;
    if (!serverMissing) throw new Error(data?.error || `Erro HTTP ${res.status} na API do Gemini`);
    proxyError = new Error('Servidor de IA indisponível.');
  } catch (e: any) {
    const offline = /failed to fetch|networkerror|load failed/i.test(e?.message || '');
    if (!offline && !proxyError) throw e;
    proxyError = proxyError || e;
  }

  if (customKey) return directGeminiGenerate(customKey, model, contents, config);

  if (typeof window !== 'undefined') window.dispatchEvent(new CustomEvent('open-gemini-tutorial-modal'));
  throw new Error('Não foi possível falar com o servidor de IA e não há chave própria do Gemini salva em Configurações.');
}

/**
 * Testa uma chave. Sem chave informada, testa a chave do sistema (GEMINI_API_KEY no servidor).
 */
export async function testGeminiKey(key?: string): Promise<{ ok: boolean; message: string }> {
  const k = cleanApiKey(key || '');
  const headers: Record<string, string> = { ...(await geminiAuthHeaders()) };
  if (k) headers['x-custom-api-key'] = k;
  try {
    const res = await fetch(apiUrl('/api/gemini/test'), { method: 'POST', headers });
    const isJson = (res.headers.get('content-type') || '').includes('application/json');
    if (isJson && res.status !== 404) {
      const data: any = await res.json().catch(() => ({}));
      return data?.ok ? { ok: true, message: data.message || 'Conexão ativa!' } : { ok: false, message: data?.error || `Erro HTTP ${res.status}` };
    }
  } catch {
    /* servidor fora do ar: tenta direto abaixo */
  }
  if (!k) return { ok: false, message: 'Servidor de IA indisponível para testar a chave do sistema.' };
  try {
    const r = await directGeminiGenerate(k, GEMINI_PRIMARY_MODEL, "Responda apenas 'OK'.");
    return { ok: true, message: `Conexão ativa! O modelo ${r.model} respondeu com sucesso.` };
  } catch (e: any) {
    return { ok: false, message: e?.message || String(e) };
  }
}
