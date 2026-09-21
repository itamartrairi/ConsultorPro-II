import type { Config } from '@netlify/functions';
import { processKiwifyWebhook } from '../../src/server/kiwifyWebhook';

/**
 * Webhook da Kiwify no Netlify: libera/bloqueia a licença em empresas_credenciadas.
 * URL para cadastrar na Kiwify:
 *   https://SEU-SITE.netlify.app/api/webhooks/kiwify?token=SEU_TOKEN_SECRETO
 * Variáveis (escopo Functions): KIWIFY_WEBHOOK_TOKEN e FIREBASE_SERVICE_ACCOUNT_JSON.
 */
export default async (req: Request): Promise<Response> => {
  if (req.method !== 'POST') {
    return new Response(JSON.stringify({ ok: false, error: 'Use POST.' }), {
      status: 405,
      headers: { 'Content-Type': 'application/json' },
    });
  }
  const url = new URL(req.url);
  const rawBody = await req.text();
  let body: any = {};
  try {
    body = rawBody ? JSON.parse(rawBody) : {};
  } catch {
    body = Object.fromEntries(new URLSearchParams(rawBody));
  }

  const result = await processKiwifyWebhook({
    token: url.searchParams.get('token') ?? undefined,
    signature: url.searchParams.get('signature') ?? undefined,
    rawBody,
    body,
  });
  return new Response(JSON.stringify(result.body), {
    status: result.status,
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
  });
};

export const config: Config = {
  path: '/api/webhooks/kiwify',
};
