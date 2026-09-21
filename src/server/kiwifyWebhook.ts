import type { Request, Response } from 'express';
import admin from 'firebase-admin';
import { timingSafeEqual, createHmac } from 'crypto';
import { getAdminDb } from './firebaseAdmin';

/**
 * ------------------------------------------------------------------------------------
 * INTEGRAÇÃO KIWIFY — Liberação automática de plano após pagamento aprovado
 * ------------------------------------------------------------------------------------
 *
 * Como configurar (uma única vez, no painel da Kiwify):
 *   1. Acesse Apps → Webhooks → Criar Webhook.
 *   2. URL do Webhook (funciona no Netlify e no server.ts):
 *        https://SEU-DOMINIO/api/webhooks/kiwify?token=SEU_TOKEN_SECRETO
 *      (defina SEU_TOKEN_SECRETO livremente e coloque o mesmo valor na variável de
 *       ambiente KIWIFY_WEBHOOK_TOKEN do servidor)
 *   3. Produtos: selecione "Todos os produtos" (ou os produtos Mensal/Anual).
 *   4. Eventos: marque pelo menos "Compra aprovada" (compra_aprovada). Recomenda-se
 *      marcar também "Assinatura Renovada", "Assinatura Cancelada", "Assinatura
 *      Atrasada", "Compra reembolsada" e "Chargeback" para manter o status em dia.
 *   5. Salve e clique em "Testar Webhook" — confira nos logs o formato exato do
 *      payload enviado pela sua conta (alguns campos podem variar por conta/plano
 *      da Kiwify). Se necessário, ajuste os nomes de campo abaixo para bater com o
 *      que aparece no log de teste.
 *
 * Variáveis de ambiente necessárias no servidor:
 *   - KIWIFY_WEBHOOK_TOKEN            → o token secreto escolhido acima (OBRIGATÓRIO:
 *                                       sem ele o webhook recusa todas as chamadas)
 *   - FIREBASE_SERVICE_ACCOUNT_JSON   → o conteúdo JSON de uma Service Account do
 *                                       Firebase (Console → Configurações do Projeto →
 *                                       Contas de Serviço → Gerar nova chave privada),
 *                                       colado como uma única linha/string.
 *
 * O que este handler faz:
 *   - Confirma que a compra foi aprovada.
 *   - Identifica o plano comprado (Mensal ou Anual) pelo valor pago.
 *   - Localiza o cadastro do consultor em `empresas_credenciadas` pelo e-mail do
 *     comprador (mesmo e-mail usado no login do sistema).
 *   - Ativa/renova a licença automaticamente (`tipoPlano`, `status`, `dataCadastro`).
 *   - Em caso de reembolso/chargeback/cancelamento, bloqueia o acesso.
 * ------------------------------------------------------------------------------------
 */

// Tenta extrair um valor de várias possíveis localizações no payload, já que o
// formato exato pode variar um pouco entre contas/versões da Kiwify.
function pick(obj: any, paths: string[]): any {
  for (const path of paths) {
    const value = path.split('.').reduce((acc, key) => (acc == null ? undefined : acc[key]), obj);
    if (value !== undefined && value !== null && value !== '') return value;
  }
  return undefined;
}

function eventStatus(body: any): string {
  return String(
    pick(body, ['webhook_event_type', 'order_status', 'Subscription.status', 'status', 'event']) || ''
  ).toLowerCase();
}

// Verificado ANTES da aprovação: um evento de cancelamento/reembolso pode vir com
// order_status "paid" do pedido original, e não pode reativar o acesso.
function normalizeCanceledStatus(body: any): boolean {
  const status = eventStatus(body);
  return ['refunded', 'refund', 'chargeback', 'canceled', 'cancelled', 'reembolsad', 'atrasad', 'late', 'inactive'].some((s) =>
    status.includes(s)
  );
}

function normalizeApprovedStatus(body: any): boolean {
  const status = eventStatus(body);
  // Compara por palavra inteira para que "inactive" não seja lido como "active".
  // Ex.: "order_approved" → ["order", "approved"]; "compra_aprovada" → ["compra", "aprovada"].
  const words = status.split(/[^a-z]+/).filter(Boolean);
  return words.some((w) => ['paid', 'approved', 'aprovada', 'aprovado', 'active', 'renewed'].includes(w));
}

function tokenMatches(provided: unknown, expected: string): boolean {
  if (typeof provided !== 'string') return false;
  const a = Buffer.from(provided);
  const b = Buffer.from(expected);
  return a.length === b.length && timingSafeEqual(a, b);
}

function resolvePlano(body: any): 'Mensal' | 'Anual' {
  const price = Number(
    pick(body, ['Commissions.charge_amount', 'charge_amount', 'total_price', 'Subscription.charge_amount', 'amount']) ?? 0
  );
  const productName = String(pick(body, ['Product.product_name', 'product_name']) || '').toLowerCase();
  if (productName.includes('anual') || price >= 200) return 'Anual';
  return 'Mensal';
}

export interface WebhookResult {
  status: number;
  body: any;
}

/**
 * Núcleo do webhook, independente do servidor (Express ou Netlify Function).
 * Autenticação: ?token=KIWIFY_WEBHOOK_TOKEN na URL, ou assinatura HMAC-SHA1 do corpo
 * (?signature=...) calculada com o mesmo token, quando o corpo bruto estiver disponível.
 */
export async function processKiwifyWebhook(input: {
  token?: unknown;
  signature?: unknown;
  rawBody?: string;
  body: any;
}): Promise<WebhookResult> {
  const res = {
    _status: 200,
    status(code: number) { this._status = code; return this; },
    json(payload: any): WebhookResult { return { status: this._status, body: payload }; },
  };
  try {
    const expectedToken = process.env.KIWIFY_WEBHOOK_TOKEN;
    const providedToken = input.token;
    if (!expectedToken) {
      // Sem token configurado, qualquer pessoa poderia liberar licenças chamando esta URL.
      console.error('[Kiwify Webhook] KIWIFY_WEBHOOK_TOKEN não configurada — webhook recusado.');
      return res.status(503).json({ ok: false, error: 'Webhook não configurado no servidor.' });
    }
    const signatureOk =
      typeof input.signature === 'string' &&
      typeof input.rawBody === 'string' &&
      tokenMatches(input.signature, createHmac('sha1', expectedToken).update(input.rawBody).digest('hex'));
    if (!signatureOk && !tokenMatches(providedToken, expectedToken)) {
      console.warn('[Kiwify Webhook] Token inválido recebido.');
      return res.status(401).json({ ok: false, error: 'Token inválido.' });
    }

    const body = input.body || {};
    const email = String(pick(body, ['Customer.email', 'customer.email', 'email']) || '').trim().toLowerCase();
    const nome = pick(body, ['Customer.full_name', 'customer.full_name', 'full_name']);
    const orderId = pick(body, ['order_id', 'id']);

    if (!email) {
      console.warn('[Kiwify Webhook] Payload recebido sem e-mail do comprador. Ignorando.', JSON.stringify(body).slice(0, 500));
      return res.status(200).json({ ok: true, ignored: true, reason: 'Sem e-mail do comprador no payload.' });
    }

    const db = getAdminDb();
    if (!db) {
      return res.status(500).json({ ok: false, error: 'Firebase Admin não configurado no servidor.' });
    }

    // Busca o cadastro do consultor pelo e-mail (mesmo e-mail do login no sistema)
    const snap = await db.collection('empresas_credenciadas').where('email', '==', email).limit(1).get();

    if (snap.empty) {
      console.warn(`[Kiwify Webhook] Nenhum cadastro encontrado para o e-mail ${email}. A compra foi registrada, mas precisa de vinculação manual.`);
      return res.status(200).json({ ok: true, matched: false, email });
    }

    const docRef = snap.docs[0].ref;

    if (normalizeCanceledStatus(body)) {
      await docRef.update({
        status: 'Bloqueada',
        kiwifyOrderId: orderId || null,
        dataCancelamentoKiwify: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: new Date().toISOString(),
      });
      console.log(`[Kiwify Webhook] Acesso bloqueado para ${email} (evento de cancelamento/reembolso).`);
      return res.status(200).json({ ok: true, matched: true, blocked: true });
    }

    if (normalizeApprovedStatus(body)) {
      const plano = resolvePlano(body);
      await docRef.update({
        tipoPlano: plano,
        status: 'Ativa',
        dataCadastro: admin.firestore.FieldValue.serverTimestamp(), // reinicia o período de vigência
        kiwifyOrderId: orderId || null,
        kiwifyNomeComprador: nome || null,
        dataAtivacaoKiwify: admin.firestore.FieldValue.serverTimestamp(),
        updatedAt: new Date().toISOString(),
      });
      console.log(`[Kiwify Webhook] Plano ${plano} ativado para ${email} (pedido ${orderId}).`);
      return res.status(200).json({ ok: true, matched: true, plano });
    }

    // Outros eventos (boleto gerado, pix gerado, carrinho abandonado, etc.) são apenas confirmados.
    return res.status(200).json({ ok: true, matched: true, ignoredEvent: true });
  } catch (error: any) {
    console.error('[Kiwify Webhook] Erro inesperado:', error);
    return res.status(500).json({ ok: false, error: error?.message || 'Erro inesperado.' });
  }
}

// Adaptador Express (server.ts).
export async function handleKiwifyWebhook(req: Request, res: Response) {
  const result = await processKiwifyWebhook({
    token: req.query.token,
    signature: req.query.signature,
    body: req.body,
  });
  return res.status(result.status).json(result.body);
}
