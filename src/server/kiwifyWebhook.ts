import type { Request, Response } from 'express';
import admin from 'firebase-admin';

/**
 * ------------------------------------------------------------------------------------
 * INTEGRAÇÃO KIWIFY — Liberação automática de plano após pagamento aprovado
 * ------------------------------------------------------------------------------------
 *
 * Como configurar (uma única vez, no painel da Kiwify):
 *   1. Acesse Apps → Webhooks → Criar Webhook.
 *   2. URL do Webhook:
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
 *   - KIWIFY_WEBHOOK_TOKEN            → o token secreto escolhido acima
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

let adminApp: admin.app.App | null = null;

function getAdminApp(): admin.app.App | null {
  if (adminApp) return adminApp;
  const raw = process.env.FIREBASE_SERVICE_ACCOUNT_JSON;
  if (!raw) {
    console.error('[Kiwify Webhook] FIREBASE_SERVICE_ACCOUNT_JSON não configurada — não é possível acessar o Firestore.');
    return null;
  }
  try {
    const serviceAccount = JSON.parse(raw);
    adminApp = admin.apps.length
      ? (admin.app() as admin.app.App)
      : admin.initializeApp({ credential: admin.credential.cert(serviceAccount) });
    return adminApp;
  } catch (e) {
    console.error('[Kiwify Webhook] Falha ao inicializar o Firebase Admin:', e);
    return null;
  }
}

// Tenta extrair um valor de várias possíveis localizações no payload, já que o
// formato exato pode variar um pouco entre contas/versões da Kiwify.
function pick(obj: any, paths: string[]): any {
  for (const path of paths) {
    const value = path.split('.').reduce((acc, key) => (acc == null ? undefined : acc[key]), obj);
    if (value !== undefined && value !== null && value !== '') return value;
  }
  return undefined;
}

function normalizeApprovedStatus(body: any): boolean {
  const status = String(
    pick(body, ['order_status', 'webhook_event_type', 'Subscription.status', 'status', 'event']) || ''
  ).toLowerCase();
  return ['paid', 'approved', 'compra_aprovada', 'aprovado', 'active'].some((s) => status.includes(s));
}

function normalizeCanceledStatus(body: any): boolean {
  const status = String(
    pick(body, ['order_status', 'webhook_event_type', 'Subscription.status', 'status', 'event']) || ''
  ).toLowerCase();
  return ['refunded', 'refund', 'chargeback', 'canceled', 'cancelled', 'reembolsad', 'atrasad', 'late'].some((s) =>
    status.includes(s)
  );
}

function resolvePlano(body: any): 'Mensal' | 'Anual' {
  const price = Number(
    pick(body, ['Commissions.charge_amount', 'charge_amount', 'total_price', 'Subscription.charge_amount', 'amount']) ?? 0
  );
  const productName = String(pick(body, ['Product.product_name', 'product_name']) || '').toLowerCase();
  if (productName.includes('anual') || price >= 200) return 'Anual';
  return 'Mensal';
}

export async function handleKiwifyWebhook(req: Request, res: Response) {
  try {
    const expectedToken = process.env.KIWIFY_WEBHOOK_TOKEN;
    const providedToken = req.query.token;
    if (expectedToken && providedToken !== expectedToken) {
      console.warn('[Kiwify Webhook] Token inválido recebido.');
      return res.status(401).json({ ok: false, error: 'Token inválido.' });
    }

    const body = req.body || {};
    const email = String(pick(body, ['Customer.email', 'customer.email', 'email']) || '').trim().toLowerCase();
    const nome = pick(body, ['Customer.full_name', 'customer.full_name', 'full_name']);
    const orderId = pick(body, ['order_id', 'id']);

    if (!email) {
      console.warn('[Kiwify Webhook] Payload recebido sem e-mail do comprador. Ignorando.', JSON.stringify(body).slice(0, 500));
      return res.status(200).json({ ok: true, ignored: true, reason: 'Sem e-mail do comprador no payload.' });
    }

    const app = getAdminApp();
    if (!app) {
      return res.status(500).json({ ok: false, error: 'Firebase Admin não configurado no servidor.' });
    }
    const db = app.firestore();

    // Busca o cadastro do consultor pelo e-mail (mesmo e-mail do login no sistema)
    const snap = await db.collection('empresas_credenciadas').where('email', '==', email).limit(1).get();

    if (snap.empty) {
      console.warn(`[Kiwify Webhook] Nenhum cadastro encontrado para o e-mail ${email}. A compra foi registrada, mas precisa de vinculação manual.`);
      return res.status(200).json({ ok: true, matched: false, email });
    }

    const docRef = snap.docs[0].ref;

    if (normalizeApprovedStatus(body)) {
      const plano = resolvePlano(body);
      await docRef.update({
        tipoPlano: plano,
        status: 'Ativa',
        dataCadastro: admin.firestore.FieldValue.serverTimestamp(), // reinicia o período de vigência
        kiwifyOrderId: orderId || null,
        kiwifyNomeComprador: nome || null,
        dataAtivacaoKiwify: admin.firestore.FieldValue.serverTimestamp(),
      });
      console.log(`[Kiwify Webhook] Plano ${plano} ativado para ${email} (pedido ${orderId}).`);
      return res.status(200).json({ ok: true, matched: true, plano });
    }

    if (normalizeCanceledStatus(body)) {
      await docRef.update({
        status: 'Bloqueada',
        kiwifyOrderId: orderId || null,
        dataCancelamentoKiwify: admin.firestore.FieldValue.serverTimestamp(),
      });
      console.log(`[Kiwify Webhook] Acesso bloqueado para ${email} (evento de cancelamento/reembolso).`);
      return res.status(200).json({ ok: true, matched: true, blocked: true });
    }

    // Outros eventos (boleto gerado, pix gerado, carrinho abandonado, etc.) são apenas confirmados.
    return res.status(200).json({ ok: true, matched: true, ignoredEvent: true });
  } catch (error: any) {
    console.error('[Kiwify Webhook] Erro inesperado:', error);
    return res.status(500).json({ ok: false, error: error?.message || 'Erro inesperado.' });
  }
}
