import assert from 'node:assert/strict';
import {
  analisarDiagnostico, montarAtividades, estruturarPlano, enriquecerComIA, montarTextosRelatorio,
  isPerguntaNegativa, papelDaAtividade, horasDe, MAX_ATIVIDADES,
  ATIVIDADE_DIAGNOSTICO, ATIVIDADE_DEVOLUTIVA, ATIVIDADE_FERRAMENTAS, ATIVIDADE_ANALISE_FINAL, ATIVIDADE_RELATORIO_FINAL,
} from '../src/lib/domain/planoAcao';

const r = (idProblema: string, pergunta: string, resposta: any, peso = 1, area = 'Finanças') =>
  ({ id: Math.random().toString(), diagnosticoId: 'd1', premissaId: 'p', idProblema, problema: idProblema, pergunta, peso, area, resposta, observacao: '', score: 0 }) as any;
const prob = (id: string, area: string, impacto: string) => ({ id, descricao_problemas: `Problema ${id}`, area, impacto }) as any;

function verificaEstrutura(lista: any[], carga: number, rotulo: string) {
  assert.equal(lista[0].nome, ATIVIDADE_DIAGNOSTICO, `${rotulo}: 1ª = diagnóstico`);
  assert.equal(lista[1].nome, ATIVIDADE_DEVOLUTIVA, `${rotulo}: 2ª = devolutiva`);
  assert.equal(lista[2].nome, ATIVIDADE_FERRAMENTAS, `${rotulo}: 3ª = ferramentas`);
  assert.equal(lista[lista.length - 2].nome, ATIVIDADE_ANALISE_FINAL, `${rotulo}: penúltima = análise final`);
  assert.equal(lista[lista.length - 2].cargaHoraria, '2h', `${rotulo}: análise final com 2h`);
  assert.equal(lista[lista.length - 1].nome, ATIVIDADE_RELATORIO_FINAL, `${rotulo}: última = relatório`);
  assert.ok(lista.length <= MAX_ATIVIDADES, `${rotulo}: no máximo 10 (${lista.length})`);
  assert.equal(lista.reduce((s, a) => s + horasDe(a), 0), carga, `${rotulo}: soma das horas = ${carga}`);
  lista.forEach((a, i) => assert.ok(horasDe(a) >= 2, `${rotulo}: atividade ${i + 1} com ${a.cargaHoraria}`));
}

// Semântica das perguntas
assert.equal(isPerguntaNegativa('O produtor já tentou acessar crédito e foi rejeitado?'), true);
assert.equal(isPerguntaNegativa('Possui dívidas em atraso?'), true);
assert.equal(isPerguntaNegativa('Possui controle financeiro?'), false);
assert.equal(isPerguntaNegativa('Tem controle de estoque para evitar perdas?'), false);
console.log('  OK: semântica das perguntas');

// Classificação
const respostas = [
  r('A', 'Possui controle de fluxo de caixa?', 'Não', 3),
  r('A', 'Separa contas pessoais e da empresa?', 'Não', 2),
  r('B', 'Possui dívidas em atraso?', 'Sim', 2, 'Crédito'),
  r('C', 'Tem controle de estoque?', 'Parcial', 1, 'Estoque'),
  r('D', 'Possui cadastro de clientes?', 'Sim', 1, 'Comercial'),
  r('E', 'Calcula o custo de produção?', 'Não', 1, 'Custos'),
];
const problemas = [prob('A', 'Finanças', 'Alto'), prob('B', 'Crédito', 'Médio'), prob('C', 'Estoque', 'Baixo'), prob('D', 'Comercial', 'Alto'), prob('E', 'Custos', 'Médio')];
const solucoes = [{ id: 's1', idProblema: 'A', problema: 'Problema A', area: 'Finanças', solucao_recomendada: 'Implantar fluxo de caixa', acoes_sugeridas: 'Registrar entradas e saídas diariamente', kpis_sugeridos: '100% dos lançamentos', resultado_esperado: 'Fluxo de caixa em uso', responsavel_sugerido: 'Cliente', prazo_sugerido: '', comentario_sucesso: '' }] as any;
const an = analisarDiagnostico(respostas, problemas, solucoes);
assert.equal(an.totalRespondidas, 6);
assert.deepEqual(an.problemas.map((p) => p.idProblema), ['A', 'B', 'E', 'C'], 'ordem por criticidade');
assert.equal(an.problemas[0].nivel, 'Crítico');
assert.equal(an.problemas[1].nivel, 'Crítico', '"Sim" em pergunta negativa é problema');
assert.equal(an.problemas[3].nivel, 'Moderado');
assert.deepEqual(an.pontosFortes.map((p) => p.problema), ['Problema D']);
assert.equal(an.problemas[0].solucao?.id, 's1');
console.log('  OK: classificação dos problemas críticos');

// Plano padrão 34h
const plano = montarAtividades(an, '34h');
verificaEstrutura(plano, 34, '34h');
assert.deepEqual(plano.slice(3, -2).map((a) => a.idProblema), ['A', 'B', 'E', 'C'], 'problemas no meio, em ordem');
assert.match(plano[3].nome, /fluxo de caixa/i, 'usa a solução cadastrada');
assert.match(plano[2].descricao, /fluxo de caixa|custos/i, 'ferramentas conforme as áreas');
console.log('  OK: plano 34h —', plano.map((a) => `${a.cargaHoraria}`).join(' + '));

// Muitos problemas → no máximo 10 atividades
const muitos = Array.from({ length: 12 }, (_, i) => r(`P${i}`, `Possui controle ${i}?`, 'Não', 1));
const anMuitos = analisarDiagnostico(muitos, [], []);
for (const carga of [34, 40, 60, 16, 14]) {
  const p = montarAtividades(anMuitos, carga);
  verificaEstrutura(p, carga, `${carga}h/12 problemas`);
  assert.ok(p.some((a) => /complementares/.test(a.nome)), `${carga}h: excedentes agrupados`);
}
console.log('  OK: até 10 atividades com agrupamento (14, 16, 34, 40 e 60h)');

// Sem lacunas
const semLacuna = montarAtividades(analisarDiagnostico([r('X', 'Possui controle?', 'Sim')], [], []), 20);
verificaEstrutura(semLacuna, 20, 'sem lacunas');
console.log('  OK: diagnóstico sem lacunas');

// Estrutura antiga (replicação/modelos antigos) → nova estrutura
const antigo = [
  { nome: 'Entendimento da demanda e diagnóstico inicial', cargaHoraria: '2h' },
  { nome: 'Controle de custos', cargaHoraria: '4h' },
  { nome: 'Desenvolvimento de sistema, aplicativos ou planilhas', cargaHoraria: '8h' },
  { nome: 'Análise final de atividades da consultoria', cargaHoraria: '2h' },
  { nome: 'Relatório final e encerramento', cargaHoraria: '2h' },
  { nome: 'Precificação', cargaHoraria: '4h' },
] as any[];
const conv = estruturarPlano(antigo, 30);
verificaEstrutura(conv, 30, 'conversão');
assert.deepEqual(conv.slice(3, -2).map((a) => a.nome), ['Controle de custos', 'Precificação']);
assert.equal(conv.filter((a) => a.nome === ATIVIDADE_ANALISE_FINAL).length, 1, 'mantém uma única análise final');
assert.equal(papelDaAtividade('Relatório final e encerramento'), 'relatorio');
console.log('  OK: conversão de planos/modelos antigos');

// IA: melhora textos, mas não quebra a estrutura
const fakeAI = { models: { generateContent: async () => ({ text: JSON.stringify([
  { idProblema: 'A', nome: 'Implantar fluxo de caixa diário', descricao: 'd', solucaoProposta: 's', resultadoEsperado: '100% dos lançamentos em 30 dias' },
  { idProblema: 'B', nome: 'Relatório final e encerramento', descricao: 'tentativa de virar atividade fixa', solucaoProposta: 's', resultadoEsperado: 'r' },
  { idProblema: '__ferramentas__', nome: 'x', descricao: 'Planilha de fluxo de caixa e controle de dívidas', solucaoProposta: 's', resultadoEsperado: '2 planilhas' },
  { idProblema: 'NAO_EXISTE', nome: 'Inventada', descricao: 'd', solucaoProposta: 's', resultadoEsperado: 'r' },
]) }) } };
const ia = await enriquecerComIA(fakeAI, plano, an);
assert.equal(ia.usouIA, true);
verificaEstrutura(ia.atividades, 34, 'após IA');
assert.equal(ia.atividades[3].nome, 'Implantar fluxo de caixa diário');
assert.equal(ia.atividades[4].nome, plano[4].nome, 'IA não pode renomear para atividade fixa');
assert.equal(ia.atividades[2].descricao, 'Planilha de fluxo de caixa e controle de dívidas');
assert.equal(ia.atividades.length, plano.length, 'IA não adiciona atividades');
const falha = await enriquecerComIA({ models: { generateContent: async () => { throw new Error('401'); } } }, plano, an);
assert.equal(falha.usouIA, false); assert.equal(falha.atividades, plano);
console.log('  OK: IA só altera textos; falha da IA mantém o plano');

// Relatório
const rel = montarTextosRelatorio(an, plano, { nomeEmpresa: 'Empresa X' });
assert.match(rel.objetivo, /^Solucionar os 3 problemas críticos identificados no diagnóstico da empresa Empresa X/);
console.log('   Objetivo:', rel.objetivo);
assert.match(rel.solucoesIndicadas, /PROBLEMAS IDENTIFICADOS:[\s\S]*\[Crítico\] Problema A[\s\S]*SOLUÇÕES \/ AÇÕES PROPOSTAS:[\s\S]*PONTOS FORTES:/);
console.log('  OK: textos do relatório');
console.log('Todos os testes de planoAcao passaram.');
