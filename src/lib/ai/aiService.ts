/**
 * Serviço Multi-IA para ConsultorPro-II.
 * Suporta Groq Cloud (Llama 3.3 70B Versatile) e Google Gemini (2.5 Flash).
 */

export type AiProvider = 'groq' | 'gemini';

const GROQ_DEFAULT_MODEL = 'llama-3.3-70b-versatile';
const GROQ_API_URL = 'https://api.groq.com/openai/v1/chat/completions';

export function getActiveAiProvider(): AiProvider {
  try {
    const saved = localStorage.getItem('active_ai_provider');
    if (saved === 'gemini' || saved === 'groq') return saved;
    // Se o usuário tem chave da Groq configurada, prioriza Groq por padrão
    if (getGroqApiKey()) return 'groq';
    return 'groq'; // Padrão recomendado
  } catch {
    return 'groq';
  }
}

export function setActiveAiProvider(provider: AiProvider): void {
  try {
    localStorage.setItem('active_ai_provider', provider);
  } catch {
    /* ignore */
  }
}

export function isValidGroqApiKey(k: any): boolean {
  if (!k || typeof k !== 'string') return false;
  const t = k.trim();
  return t.startsWith('gsk_') && t.length >= 25;
}

export function getGroqApiKey(): string {
  try {
    const stored = (localStorage.getItem('custom_groq_api_key') || '').trim();
    if (isValidGroqApiKey(stored)) return stored;
    const envKey = (
      (import.meta as any).env?.VITE_GROQ_API_KEY ||
      (import.meta as any).env?.GROQ_API_KEY ||
      (typeof process !== 'undefined' ? (process.env?.VITE_GROQ_API_KEY || process.env?.GROQ_API_KEY) : '') ||
      ''
    ).trim();
    if (isValidGroqApiKey(envKey)) return envKey;
    return stored; // Retorna chave salva mesmo se formato for atípico
  } catch {
    return '';
  }
}

export function saveGroqApiKey(key: string): void {
  try {
    const clean = key.trim();
    if (!clean) {
      localStorage.removeItem('custom_groq_api_key');
    } else {
      localStorage.setItem('custom_groq_api_key', clean);
    }
  } catch {
    /* ignore */
  }
}

export async function testGroqKey(keyToTest: string): Promise<{ ok: boolean; message: string; model?: string }> {
  const cleanKey = keyToTest.trim();
  if (!cleanKey) {
    return { ok: false, message: 'Por favor, informe a chave de API da Groq (inicia com gsk_).' };
  }

  try {
    const response = await fetch(GROQ_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${cleanKey}`
      },
      body: JSON.stringify({
        model: GROQ_DEFAULT_MODEL,
        messages: [{ role: 'user', content: "Responda apenas 'OK'." }],
        temperature: 0.1,
        max_tokens: 10
      })
    });

    const data = await response.json();
    if (!response.ok) {
      const errMsg = data?.error?.message || `Erro HTTP ${response.status}`;
      return { ok: false, message: `Falha na Groq: ${errMsg}` };
    }

    return {
      ok: true,
      model: GROQ_DEFAULT_MODEL,
      message: `Conexão bem-sucedida com Groq (${GROQ_DEFAULT_MODEL})!`
    };
  } catch (err: any) {
    return { ok: false, message: `Erro de rede ao conectar à Groq: ${err?.message || String(err)}` };
  }
}

export interface UniversalGenerateOptions {
  contents: any;
  config?: any;
  model?: string;
}

export async function callGroqChat(options: UniversalGenerateOptions): Promise<{ text: string }> {
  const apiKey = getGroqApiKey();
  if (!apiKey) {
    throw new Error(
      "Chave da API Groq não configurada.\n\nVocê pode gerar uma chave 100% gratuita sem cartão em: https://console.groq.com/keys e inseri-la nas Configurações do sistema."
    );
  }

  let userPrompt = '';
  if (typeof options.contents === 'string') {
    userPrompt = options.contents;
  } else if (Array.isArray(options.contents)) {
    userPrompt = options.contents.map(c => typeof c === 'string' ? c : c?.parts?.map((p: any) => p.text).join('\n') || JSON.stringify(c)).join('\n');
  } else if (options.contents?.parts) {
    userPrompt = options.contents.parts.map((p: any) => p.text).join('\n');
  } else {
    userPrompt = JSON.stringify(options.contents);
  }

  const isJsonExpected =
    options.config?.responseMimeType === 'application/json' ||
    userPrompt.toLowerCase().includes('formato json') ||
    userPrompt.toLowerCase().includes('retorne em json') ||
    userPrompt.toLowerCase().includes('em json');

  const systemInstructions = isJsonExpected
    ? "Você é um consultor empresarial sênior do SEBRAE. Quando for solicitado formato JSON, responda ESTRITAMENTE em formato JSON válido, sem texto introdutório e sem blocos de código com crases."
    : "Você é um consultor empresarial sênior experiente. Forneça respostas executivas, estruturadas e práticas em Português do Brasil.";

  const bodyPayload: any = {
    model: GROQ_DEFAULT_MODEL,
    messages: [
      { role: 'system', content: systemInstructions },
      { role: 'user', content: userPrompt }
    ],
    temperature: isJsonExpected ? 0.1 : 0.4
  };

  if (isJsonExpected) {
    bodyPayload.response_format = { type: 'json_object' };
  }

  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 25000);

  try {
    const response = await fetch(GROQ_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      signal: controller.signal,
      body: JSON.stringify(bodyPayload)
    });

    clearTimeout(timeoutId);

    const data = await response.json();
    if (!response.ok) {
      const errDetail = data?.error?.message || `Erro HTTP ${response.status}`;
      throw new Error(`[Groq Error] ${errDetail}`);
    }

    const text = data?.choices?.[0]?.message?.content || '';
    if (!text) {
      throw new Error("A Groq retornou uma resposta em branco.");
    }

    return { text: text.trim() };
  } catch (err: any) {
    clearTimeout(timeoutId);
    if (err.name === 'AbortError') {
      throw new Error("Tempo limite excedido ao consultar o servidor da Groq.");
    }
    throw err;
  }
}
