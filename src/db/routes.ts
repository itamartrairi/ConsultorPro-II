import { Router } from 'express';
import { getDatabase, initPostgresTables } from './index';
import { empresas, diagnosticos, respostas, tarefasPlano, empresasCredenciadas, appMetadata } from './schema';
import { eq } from 'drizzle-orm';
import type { Request, Response, NextFunction } from 'express';
import { verifyFirebaseUser } from '../server/firebaseAdmin';

// Identidade verificada no servidor a partir do token do Firebase Auth.
// Antes, userId e isAdmin vinham do corpo da requisição — qualquer pessoa podia
// enviar isAdmin: true e baixar os dados de todos os consultores.
async function requireFirebaseUser(req: Request, res: Response, next: NextFunction) {
  const user = await verifyFirebaseUser(req.headers.authorization);
  if (!user) {
    return res.status(401).json({ ok: false, error: 'Login necessário ou sessão expirada. Entre novamente.' });
  }
  (res.locals as any).uid = user.uid;
  (res.locals as any).isAdmin = user.isAdmin;
  next();
}

export const dbRouter = Router();

// Test & Probe Netlify Database connection
dbRouter.get('/status', requireFirebaseUser, async (req, res) => {
  const dbClient = getDatabase();
  const hasEnv = !!(process.env.NETLIFY_DATABASE_URL || process.env.DATABASE_URL || process.env.POSTGRES_URL);

  if (!hasEnv) {
    return res.json({
      ok: false,
      configured: false,
      message: 'Variável de conexão do Netlify Database (NETLIFY_DATABASE_URL ou DATABASE_URL) não configurada.'
    });
  }

  if (!dbClient) {
    return res.json({
      ok: false,
      configured: true,
      message: 'Falha ao instanciar cliente do Netlify Database.'
    });
  }

  try {
    const initOk = await initPostgresTables();
    res.json({
      ok: initOk,
      configured: true,
      provider: 'Netlify Database (PostgreSQL / Neon)',
      message: initOk 
        ? 'Conexão com Netlify Database (PostgreSQL) estabelecida e tabelas validadas!' 
        : 'Banco conectado, mas houve falha na verificação das tabelas.'
    });
  } catch (err: any) {
    res.status(500).json({
      ok: false,
      configured: true,
      message: 'Erro ao comunicar com Netlify Database: ' + (err.message || err)
    });
  }
});

// Sync data endpoint: pulls and updates records
dbRouter.post('/sync', requireFirebaseUser, async (req, res) => {
  const dbClient = getDatabase();
  if (!dbClient) {
    return res.status(400).json({
      ok: false,
      error: 'Netlify Database não configurado no servidor.'
    });
  }

  const { db, sql } = dbClient;
  const userId: string = (res.locals as any).uid;
  const isAdmin: boolean = (res.locals as any).isAdmin;
  const {
    localEmpresas = [], 
    localDiagnosticos = [], 
    localRespostas = [], 
    localTarefas = [],
    localCredenciadas = []
  } = req.body;

  // Consultor comum só grava registros em seu próprio nome; admin preserva o dono original.
  const ownerFor = (item: any): string => (isAdmin ? item.ownerId || userId : userId);

  try {
    await initPostgresTables();

    // 1. Fetch Cloud Records from Postgres
    let cloudEmpresas: any[] = [];
    let cloudDiags: any[] = [];
    let cloudResps: any[] = [];
    let cloudTasks: any[] = [];
    let cloudCreds: any[] = [];

    if (isAdmin) {
      cloudEmpresas = await db.select().from(empresas);
      cloudDiags = await db.select().from(diagnosticos);
      cloudResps = await db.select().from(respostas);
      cloudTasks = await db.select().from(tarefasPlano);
      cloudCreds = await db.select().from(empresasCredenciadas);
    } else if (userId) {
      cloudEmpresas = await db.select().from(empresas).where(eq(empresas.ownerId, userId));
      cloudDiags = await db.select().from(diagnosticos).where(eq(diagnosticos.ownerId, userId));
      cloudResps = await db.select().from(respostas).where(eq(respostas.ownerId, userId));
      cloudTasks = await db.select().from(tarefasPlano).where(eq(tarefasPlano.ownerId, userId));
      cloudCreds = await db.select().from(empresasCredenciadas).where(eq(empresasCredenciadas.ownerId, userId));
    }

    // 2. Upsert incoming local items to Postgres
    for (const emp of localEmpresas) {
      if (!emp.id) continue;
      const existing = cloudEmpresas.find(e => e.id === emp.id);
      if (!existing || (emp.updatedAt && (!existing.updatedAt || new Date(emp.updatedAt) > new Date(existing.updatedAt)))) {
        await sql`
          INSERT INTO empresas (id, owner_id, nome, cnpj, tipo_empresa, status, telefone, email, cidade, estado, created_at, updated_at, data)
          VALUES (${emp.id}, ${ownerFor(emp)}, ${emp.nome || ''}, ${emp.cnpj || ''}, ${emp.tipoEmpresa || ''}, ${emp.status || 'Ativa'}, ${emp.telefone || ''}, ${emp.email || ''}, ${emp.cidade || ''}, ${emp.estado || ''}, ${emp.createdAt || ''}, ${emp.updatedAt || ''}, ${JSON.stringify(emp)}::jsonb)
          ON CONFLICT (id) DO UPDATE SET
            nome = EXCLUDED.nome,
            cnpj = EXCLUDED.cnpj,
            tipo_empresa = EXCLUDED.tipo_empresa,
            status = EXCLUDED.status,
            telefone = EXCLUDED.telefone,
            email = EXCLUDED.email,
            cidade = EXCLUDED.cidade,
            estado = EXCLUDED.estado,
            updated_at = EXCLUDED.updated_at,
            data = EXCLUDED.data
          WHERE ${isAdmin}::boolean OR empresas.owner_id = ${userId};
        `;
      }
    }

    for (const diag of localDiagnosticos) {
      if (!diag.id) continue;
      const existing = cloudDiags.find(d => d.id === diag.id);
      if (!existing || (diag.updatedAt && (!existing.updatedAt || new Date(diag.updatedAt) > new Date(existing.updatedAt)))) {
        await sql`
          INSERT INTO diagnosticos (id, owner_id, empresa_id, tipo_empresa, data_diagnostico, status, nome_projeto, percentual_geral, areas_diagnostico, created_at, updated_at, data)
          VALUES (${diag.id}, ${ownerFor(diag)}, ${diag.empresaId || ''}, ${diag.tipoEmpresa || ''}, ${diag.dataDiagnostico || ''}, ${diag.status || ''}, ${diag.nomeProjeto || diag.nome || ''}, ${diag.percentualGeral || 0}, ${JSON.stringify(diag.areasDiagnostico || [])}::jsonb, ${diag.createdAt || ''}, ${diag.updatedAt || ''}, ${JSON.stringify(diag)}::jsonb)
          ON CONFLICT (id) DO UPDATE SET
            empresa_id = EXCLUDED.empresa_id,
            tipo_empresa = EXCLUDED.tipo_empresa,
            data_diagnostico = EXCLUDED.data_diagnostico,
            status = EXCLUDED.status,
            nome_projeto = EXCLUDED.nome_projeto,
            percentual_geral = EXCLUDED.percentual_geral,
            areas_diagnostico = EXCLUDED.areas_diagnostico,
            updated_at = EXCLUDED.updated_at,
            data = EXCLUDED.data
          WHERE ${isAdmin}::boolean OR diagnosticos.owner_id = ${userId};
        `;
      }
    }

    for (const resp of localRespostas) {
      if (!resp.id) continue;
      await sql`
        INSERT INTO respostas (id, owner_id, diagnostico_id, premissa_id, area, resposta, observacao, peso, created_at, updated_at, data)
        VALUES (${resp.id}, ${ownerFor(resp)}, ${resp.diagnosticoId || ''}, ${resp.premissaId || ''}, ${resp.area || ''}, ${resp.resposta || ''}, ${resp.observacao || ''}, ${resp.peso || 1}, ${resp.createdAt || ''}, ${resp.updatedAt || ''}, ${JSON.stringify(resp)}::jsonb)
        ON CONFLICT (id) DO UPDATE SET
          resposta = EXCLUDED.resposta,
          observacao = EXCLUDED.observacao,
          peso = EXCLUDED.peso,
          updated_at = EXCLUDED.updated_at,
          data = EXCLUDED.data
          WHERE ${isAdmin}::boolean OR respostas.owner_id = ${userId};
      `;
    }

    for (const t of localTarefas) {
      if (!t.id) continue;
      await sql`
        INSERT INTO tarefas_plano (id, owner_id, diagnostico_id, area, acao, responsavel, prazo, status, prioridade, created_at, updated_at, data)
        VALUES (${t.id}, ${ownerFor(t)}, ${t.diagnosticoId || ''}, ${t.area || ''}, ${t.acao || ''}, ${t.responsavel || ''}, ${t.prazo || ''}, ${t.status || 'Pendente'}, ${t.prioridade || 'Média'}, ${t.createdAt || ''}, ${t.updatedAt || ''}, ${JSON.stringify(t)}::jsonb)
        ON CONFLICT (id) DO UPDATE SET
          area = EXCLUDED.area,
          acao = EXCLUDED.acao,
          responsavel = EXCLUDED.responsavel,
          prazo = EXCLUDED.prazo,
          status = EXCLUDED.status,
          prioridade = EXCLUDED.prioridade,
          updated_at = EXCLUDED.updated_at,
          data = EXCLUDED.data
          WHERE ${isAdmin}::boolean OR tarefas_plano.owner_id = ${userId};
      `;
    }

    // Return fresh consolidated dataset from PostgreSQL
    const finalEmpresas = isAdmin ? await db.select().from(empresas) : await db.select().from(empresas).where(eq(empresas.ownerId, userId));
    const finalDiags = isAdmin ? await db.select().from(diagnosticos) : await db.select().from(diagnosticos).where(eq(diagnosticos.ownerId, userId));
    const finalResps = isAdmin ? await db.select().from(respostas) : await db.select().from(respostas).where(eq(respostas.ownerId, userId));
    const finalTasks = isAdmin ? await db.select().from(tarefasPlano) : await db.select().from(tarefasPlano).where(eq(tarefasPlano.ownerId, userId));
    const finalCreds = isAdmin ? await db.select().from(empresasCredenciadas) : await db.select().from(empresasCredenciadas).where(eq(empresasCredenciadas.ownerId, userId));

    res.json({
      ok: true,
      empresas: finalEmpresas.map(e => ({ ...((e.data as any) || {}), ...e })),
      diagnosticos: finalDiags.map(d => ({ ...((d.data as any) || {}), ...d })),
      respostas: finalResps.map(r => ({ ...((r.data as any) || {}), ...r })),
      tarefas: finalTasks.map(t => ({ ...((t.data as any) || {}), ...t })),
      credenciadas: finalCreds.map(c => ({ ...((c.data as any) || {}), ...c }))
    });
  } catch (err: any) {
    console.error("Erro na sincronização com Netlify Database:", err);
    res.status(500).json({ ok: false, error: err.message || err });
  }
});
