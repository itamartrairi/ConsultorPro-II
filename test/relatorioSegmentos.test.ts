import assert from 'node:assert/strict';
import { montarRelatorioSegmentos, faixaConformidade, tiposDeNegocio } from '../src/lib/domain/relatorioSegmentos';

const empresas = [
  { id: 'e1', nome: 'Fazenda Camarão Azul', tipoEmpresa: 'Carcinicultura' },
  { id: 'e2', nome: 'Aqua Norte', tipoEmpresa: 'Carcinicultura' },
  { id: 'e3', nome: 'Padaria Sol', tipoEmpresa: 'Comércio' },
  { id: 'e4', nome: 'Cliente Novo', tipoEmpresa: 'Comércio' },
  { id: 'e5', nome: 'Sem Segmento' },
] as any[];
const diagnosticos = [
  { id: 'd1a', empresaId: 'e1', dataDiagnostico: '2026-01-10' },
  { id: 'd1b', empresaId: 'e1', dataDiagnostico: '2026-06-10' }, // mais recente → usado
  { id: 'd2', empresaId: 'e2', dataDiagnostico: '2026-05-01' },
  { id: 'd3', empresaId: 'e3', dataDiagnostico: '2026-04-01' },
  { id: 'd5', empresaId: 'e5', dataDiagnostico: '2026-04-01' },
  { id: 'd4', empresaId: 'e4', dataDiagnostico: '2026-07-01' }, // sem respostas
] as any[];
const r = (diag: string, prob: string, area: string, pergunta: string, resposta: string, peso = 1) =>
  ({ id: `${diag}-${pergunta}`, diagnosticoId: diag, idProblema: prob, problema: prob, area, pergunta, resposta, peso }) as any;
const problemas = [
  { id: 'Sem fluxo de caixa', descricao_problemas: 'Sem fluxo de caixa', area: 'Finanças', impacto: 'Alto' },
  { id: 'Sem controle de ração', descricao_problemas: 'Sem controle de ração', area: 'Produção', impacto: 'Médio' },
] as any[];
const respostas = [
  r('d1a', 'Sem fluxo de caixa', 'Finanças', 'Possui fluxo de caixa?', 'Não', 3), // antigo, ignorado
  r('d1b', 'Sem fluxo de caixa', 'Finanças', 'Possui fluxo de caixa?', 'Sim', 3),
  r('d1b', 'Sem controle de ração', 'Produção', 'Controla a ração?', 'Não', 2),
  r('d2', 'Sem fluxo de caixa', 'Finanças', 'Possui fluxo de caixa?', 'Não', 3),
  r('d2', 'Sem controle de ração', 'Produção', 'Controla a ração?', 'Não', 2),
  r('d3', 'Sem fluxo de caixa', 'Finanças', 'Possui fluxo de caixa?', 'Parcial', 3),
  r('d3', 'Dívidas', 'Finanças', 'Possui dívidas em atraso?', 'Sim', 2),
  r('d5', 'Sem fluxo de caixa', 'Finanças', 'Possui fluxo de caixa?', 'Sim', 3),
];

const rel = montarRelatorioSegmentos({ empresas, diagnosticos, respostas, problemas });
assert.equal(rel.totalClientes, 5);
assert.equal(rel.totalDiagnosticados, 4);
assert.deepEqual(rel.segmentos.map((s) => s.segmento), ['Carcinicultura', 'Comércio', 'Não informado']);

const carc = rel.segmentos[0];
assert.equal(carc.clientesDiagnosticados, 2);
const azul = carc.clientes.find((c) => c.cliente === 'Fazenda Camarão Azul')!;
assert.equal(azul.diagnosticoId, 'd1b', 'usa o diagnóstico mais recente');
assert.equal(azul.totalDiagnosticos, 2);
assert.equal(azul.areas['Finanças'], 100);
assert.equal(carc.clientes[0].cliente, 'Aqua Norte', 'clientes do menor para o maior índice');
assert.equal(carc.clientes[0].conformidade, 0);
// Todos os problemas diagnosticados do cliente, com criticidade e evidências
assert.deepEqual(carc.clientes[0].problemas.map((p) => p.problema).sort(), ['Sem controle de ração', 'Sem fluxo de caixa']);
assert.ok(carc.clientes[0].problemas.every((p) => p.nivel && p.area && p.evidencias.length > 0));
assert.deepEqual(carc.clientes[0].problemas.find((p) => p.problema === 'Sem fluxo de caixa')!.evidencias, ['Possui fluxo de caixa? (Não)']);
assert.equal(azul.problemas.length, 1, 'cliente com fluxo de caixa ok tem só 1 problema');
const racao = carc.recorrentes.find((x) => x.problema === 'Sem controle de ração')!;
assert.equal(racao.clientes, 2); assert.equal(racao.percentual, 100);
assert.equal(carc.areas[0].area, 'Produção', 'área mais frágil primeiro');
assert.equal(carc.areas[0].conformidadeMedia, 0);

const com = rel.segmentos[1];
assert.deepEqual(com.clientesSemDiagnostico, ['Cliente Novo']);
assert.ok(com.recorrentes.some((x) => x.problema === 'Dívidas'), '"Sim" em pergunta negativa conta como problema');
assert.equal(rel.segmentos[2].clientes[0].conformidade, 100);
assert.equal(faixaConformidade(35).rotulo, 'Crítica');
assert.equal(faixaConformidade(85).rotulo, 'Adequada');
// Filtro por tipo de negócio (ignora acentos/maiúsculas)
const soComercio = montarRelatorioSegmentos({ empresas, diagnosticos, respostas, problemas, segmento: 'comercio' });
assert.deepEqual(soComercio.segmentos.map((s) => s.segmento), ['Comércio']);
assert.equal(soComercio.totalClientes, 2); assert.equal(soComercio.totalDiagnosticados, 1);
assert.equal(soComercio.filtroSegmento, 'comercio');
assert.equal(montarRelatorioSegmentos({ empresas, diagnosticos, respostas, segmento: 'Todos' }).segmentos.length, 3);
assert.equal(montarRelatorioSegmentos({ empresas, diagnosticos, respostas, segmento: 'Inexistente' }).totalClientes, 0);
assert.deepEqual(tiposDeNegocio(empresas, diagnosticos), ['Carcinicultura', 'Comércio', 'Não informado']);
console.log('Todos os testes de relatorioSegmentos passaram.');

// PDF de exemplo
const { gerarPdfRelatorioSegmentos } = await import('../src/lib/relatorios/pdfSegmentos');
const doc = gerarPdfRelatorioSegmentos(process.env.SO_SEGMENTO ? montarRelatorioSegmentos({ empresas, diagnosticos, respostas, problemas, segmento: process.env.SO_SEGMENTO }) : rel, { consultor: 'consultor@exemplo.com' });
const minimo = process.env.SO_SEGMENTO ? 2 : 4;
assert.ok(doc.getNumberOfPages() >= minimo, 'capa + segmentos');
if (process.env.SALVAR_PDF) (await import('node:fs')).writeFileSync(process.env.SALVAR_PDF, Buffer.from(doc.output('arraybuffer')));
