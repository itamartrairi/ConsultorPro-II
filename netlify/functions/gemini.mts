import type { Config } from '@netlify/functions';
import { geminiConfig, geminiGenerate, geminiTest, type CoreResult } from '../../src/server/geminiCore';
import { verifyFirebaseUser } from '../../src/server/firebaseAdmin';

/**
 * Proxy do Gemini no Netlify. A chave do sistema fica SOMENTE na variável de ambiente
 * GEMINI_API_KEY (Site configuration → Environment variables, escopo Functions).
 *
 *   GET  /api/gemini/config    → { hasSystemKey }
 *   POST /api/gemini/generate  → { text }            (exige login do Firebase)
 *   POST /api/gemini/test      → { ok, model }       (exige login do Firebase)
 */

// Bearer token (não cookie), então liberar a origem é seguro e permite o app desktop (Electron).
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization, x-custom-api-key',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
};

function reply(r: CoreResult): Response {
  return new Response(JSON.stringify(r.body), {
    status: r.status,
    headers: { 'Content-Type': 'application/json; charset=utf-8', ...CORS },
  });
}

export default async (req: Request): Promise<Response> => {
  if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS });

  const action = new URL(req.url).pathname.split('/').pop();

  if (action === 'config' && req.method === 'GET') return reply(geminiConfig());

  if (req.method !== 'POST' || (action !== 'generate' && action !== 'test')) {
    return reply({ status: 404, body: { error: 'Rota não encontrada.' } });
  }

  const user = await verifyFirebaseUser(req.headers.get('authorization'));
  if (!user) return reply({ status: 401, body: { ok: false, error: 'Faça login novamente para usar a IA.' } });

  const customKey = req.headers.get('x-custom-api-key');
  if (action === 'test') return reply(await geminiTest(customKey));

  let body: any;
  try {
    body = await req.json();
  } catch {
    return reply({ status: 400, body: { error: 'JSON inválido.' } });
  }
  return reply(await geminiGenerate(body, customKey));
};

export const config: Config = {
  path: ['/api/gemini/config', '/api/gemini/generate', '/api/gemini/test'],
};
