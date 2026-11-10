import type { Config } from '@netlify/functions';

/**
 * Hora do servidor, usada para contar os dias da licença sem depender do relógio do
 * computador do usuário (atrasar o relógio do PC dava dias a mais de teste).
 *
 *   GET /api/time  → { now: <milissegundos> }
 */

// Sem dados sensíveis: liberar a origem permite o app desktop (Electron).
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, OPTIONS',
};

export default async (req: Request): Promise<Response> => {
  if (req.method === 'OPTIONS') return new Response(null, { status: 204, headers: CORS });
  return new Response(JSON.stringify({ now: Date.now() }), {
    status: 200,
    headers: { ...CORS, 'Content-Type': 'application/json', 'Cache-Control': 'no-store' },
  });
};

export const config: Config = { path: '/api/time' };
