import { pgTable, text, timestamp, jsonb, boolean, integer } from 'drizzle-orm/pg-core';

export const empresas = pgTable('empresas', {
  id: text('id').primaryKey(),
  ownerId: text('owner_id').notNull().default(''),
  nome: text('nome').notNull(),
  cnpj: text('cnpj'),
  tipoEmpresa: text('tipo_empresa'),
  status: text('status').default('Ativa'),
  telefone: text('telefone'),
  email: text('email'),
  cidade: text('cidade'),
  estado: text('estado'),
  createdAt: text('created_at'),
  updatedAt: text('updated_at'),
  data: jsonb('data').default({})
});

export const diagnosticos = pgTable('diagnosticos', {
  id: text('id').primaryKey(),
  ownerId: text('owner_id').notNull().default(''),
  empresaId: text('empresa_id').notNull(),
  tipoEmpresa: text('tipo_empresa'),
  dataDiagnostico: text('data_diagnostico'),
  status: text('status'),
  nomeProjeto: text('nome_projeto'),
  percentualGeral: integer('percentual_geral').default(0),
  areasDiagnostico: jsonb('areas_diagnostico').default([]),
  createdAt: text('created_at'),
  updatedAt: text('updated_at'),
  data: jsonb('data').default({})
});

export const respostas = pgTable('respostas', {
  id: text('id').primaryKey(),
  ownerId: text('owner_id').notNull().default(''),
  diagnosticoId: text('diagnostico_id').notNull(),
  premissaId: text('premissa_id'),
  area: text('area'),
  resposta: text('resposta'),
  observacao: text('observacao'),
  peso: integer('peso').default(1),
  createdAt: text('created_at'),
  updatedAt: text('updated_at'),
  data: jsonb('data').default({})
});

export const tarefasPlano = pgTable('tarefas_plano', {
  id: text('id').primaryKey(),
  ownerId: text('owner_id').notNull().default(''),
  diagnosticoId: text('diagnostico_id').notNull(),
  area: text('area'),
  acao: text('acao'),
  responsavel: text('responsavel'),
  prazo: text('prazo'),
  status: text('status').default('Pendente'),
  prioridade: text('prioridade').default('Média'),
  createdAt: text('created_at'),
  updatedAt: text('updated_at'),
  data: jsonb('data').default({})
});

export const empresasCredenciadas = pgTable('empresas_credenciadas', {
  id: text('id').primaryKey(),
  ownerId: text('owner_id').notNull().default(''),
  razaoSocial: text('razao_social'),
  nomeFantasia: text('nome_fantasia'),
  cnpj: text('cnpj'),
  email: text('email'),
  telefone: text('telefone'),
  status: text('status').default('Ativa'),
  role: text('role').default('consultant'),
  validadeLicenca: text('validade_licenca'),
  createdAt: text('created_at'),
  updatedAt: text('updated_at'),
  data: jsonb('data').default({})
});

export const appMetadata = pgTable('app_metadata', {
  key: text('key').primaryKey(),
  value: jsonb('value'),
  updatedAt: text('updated_at')
});
