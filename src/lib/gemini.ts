import { auth } from '../firebase';

export function isValidGeminiApiKey(k: any): boolean {
  if (!k || typeof k !== 'string') return false;
  const t = k.trim();
  if (t === '' || t === 'undefined' || t === 'null') return false;
  return t.startsWith('AQ.') || t.startsWith('AIzaSy') || t.length >= 15;
}

export function readCustomGeminiKey(): string {
  try {
    const raw = (localStorage.getItem('custom_gemini_api_key') || '').trim();
    if (isValidGeminiApiKey(raw)) return raw;
    const envKey = (
      (import.meta as any).env?.VITE_GEMINI_API_KEY ||
      (import.meta as any).env?.GEMINI_API_KEY ||
      (typeof process !== 'undefined' ? (process.env?.VITE_GEMINI_API_KEY || process.env?.GEMINI_API_KEY) : '') ||
      ''
    ).trim();
    if (isValidGeminiApiKey(envKey)) return envKey;
    return '';
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
