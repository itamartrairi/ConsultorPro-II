import { neon } from '@neondatabase/serverless';
import { drizzle } from 'drizzle-orm/neon-http';
import * as schema from './schema';

export function getDatabase() {
  const connectionString = 
    process.env.NETLIFY_DATABASE_URL || 
    process.env.DATABASE_URL || 
    process.env.POSTGRES_URL;

  if (!connectionString) {
    return null;
  }

  try {
    const sql = neon(connectionString);
    const db = drizzle(sql, { schema });
    return { db, sql };
  } catch (err) {
    console.error("Error creating database client:", err);
    return null;
  }
}

export async function initPostgresTables() {
  const connectionString = 
    process.env.NETLIFY_DATABASE_URL || 
    process.env.DATABASE_URL || 
    process.env.POSTGRES_URL;

  if (!connectionString) return false;

  try {
    const sql = neon(connectionString);
    await sql`
      CREATE TABLE IF NOT EXISTS empresas (
        id TEXT PRIMARY KEY,
        owner_id TEXT DEFAULT '',
        nome TEXT NOT NULL,
        cnpj TEXT,
        tipo_empresa TEXT,
        status TEXT DEFAULT 'Ativa',
        telefone TEXT,
        email TEXT,
        cidade TEXT,
        estado TEXT,
        created_at TEXT,
        updated_at TEXT,
        data JSONB DEFAULT '{}'::jsonb
      );
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS diagnosticos (
        id TEXT PRIMARY KEY,
        owner_id TEXT DEFAULT '',
        empresa_id TEXT NOT NULL,
        tipo_empresa TEXT,
        data_diagnostico TEXT,
        status TEXT,
        nome_projeto TEXT,
        percentual_geral INTEGER DEFAULT 0,
        areas_diagnostico JSONB DEFAULT '[]'::jsonb,
        created_at TEXT,
        updated_at TEXT,
        data JSONB DEFAULT '{}'::jsonb
      );
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS respostas (
        id TEXT PRIMARY KEY,
        owner_id TEXT DEFAULT '',
        diagnostico_id TEXT NOT NULL,
        premissa_id TEXT,
        area TEXT,
        resposta TEXT,
        observacao TEXT,
        peso INTEGER DEFAULT 1,
        created_at TEXT,
        updated_at TEXT,
        data JSONB DEFAULT '{}'::jsonb
      );
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS tarefas_plano (
        id TEXT PRIMARY KEY,
        owner_id TEXT DEFAULT '',
        diagnostico_id TEXT NOT NULL,
        area TEXT,
        acao TEXT,
        responsavel TEXT,
        prazo TEXT,
        status TEXT DEFAULT 'Pendente',
        prioridade TEXT DEFAULT 'Média',
        created_at TEXT,
        updated_at TEXT,
        data JSONB DEFAULT '{}'::jsonb
      );
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS empresas_credenciadas (
        id TEXT PRIMARY KEY,
        owner_id TEXT DEFAULT '',
        razao_social TEXT,
        nome_fantasia TEXT,
        cnpj TEXT,
        email TEXT,
        telefone TEXT,
        status TEXT DEFAULT 'Ativa',
        role TEXT DEFAULT 'consultant',
        validade_licenca TEXT,
        created_at TEXT,
        updated_at TEXT,
        data JSONB DEFAULT '{}'::jsonb
      );
    `;

    await sql`
      CREATE TABLE IF NOT EXISTS app_metadata (
        key TEXT PRIMARY KEY,
        value JSONB,
        updated_at TEXT
      );
    `;

    console.log("✓ Netlify Postgres schema initialized successfully!");
    return true;
  } catch (e) {
    console.error("Failed to initialize Postgres schema:", e);
    return false;
  }
}
