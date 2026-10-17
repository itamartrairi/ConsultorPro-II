import assert from 'node:assert/strict';
import {
  analisarDiagnostico, montarAtividades, estruturarPlano, enriquecerComIA, montarTextosRelatorio,
  isPerguntaNegativa, grauLacuna, classificarStatusResposta, papelDaAtividade, horasDe, MAX_ATIVIDADES, tarefasDasAtividades, agruparPorArea,
  ATIVIDADE_DIAGNOSTICO, ATIVIDADE_DEVOLUTIVA, ATIVIDADE_FERRAMENTAS, ATIVIDADE_ANALISE_FINAL, ATIVIDADE_RELATORIO_FINAL,
  aplicarTarefaNoCronograma, removerTarefaDoCronograma, adicionarTarefaAoCronograma, sincronizarRelatorioComCronograma,
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
assert.equal(isPerguntaNegativa('O produtor comercializa apenas em feiras/festivais locais?'), true);
assert.equal(isPerguntaNegativa('Vende apenas para um cliente?'), true);
assert.equal(grauLacuna({ pergunta: 'O produtor comercializa apenas em feiras/festivais locais?', resposta: 'Não' }), 0);
assert.equal(grauLacuna({ pergunta: 'O produtor comercializa apenas em feiras/festivais locais?', resposta: 'Sim' }), 2);
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
// Uma atividade por ÁREA, das mais críticas para as menos críticas
assert.deepEqual(plano.slice(3, -2).map((a) => a.area), ['Finanças', 'Crédito', 'Custos', 'Estoque'], 'áreas em ordem de criticidade');
assert.deepEqual(plano.slice(3, -2).map((a) => a.problemasTratados!.map((t) => t.idProblema)), [['A'], ['B'], ['E'], ['C']]);
assert.match(plano[3].descricao, /Registrar entradas e saídas/i, 'usa a solução cadastrada');
assert.match(plano[3].nome, /^Finanças: (solução de 1 problema|Resolução de Problema A) \(1 crítico\)$/);
assert.match(plano[2].descricao, /fluxo de caixa|custos/i, 'ferramentas conforme as áreas');
// Áreas com críticos recebem mais horas que as moderadas
assert.ok(horasDe(plano[3]) > horasDe(plano[6]), 'área crítica com mais horas que a moderada');
// Tarefas: uma por problema nas atividades de área
const tarefas = tarefasDasAtividades(plano);
assert.equal(tarefas.length, 5 + 4, '5 atividades fixas + 1 tarefa por problema');
assert.deepEqual(tarefas.filter((t) => t.criticidade).map((t) => t.idProblema), ['A', 'B', 'E', 'C']);
assert.equal(tarefas.find((t) => t.idProblema === 'A')!.prioridade, 'Alta');
// Garantir que as ações recomendadas NÃO repetem pergunta ou resposta do diagnóstico
assert.equal(tarefas.find((t) => t.idProblema === 'A')!.acoes, 'Registrar entradas e saídas diariamente');
assert.doesNotMatch(tarefas.find((t) => t.idProblema === 'B')!.acoes, /Evidências:/i, 'não deve repetir perguntas e respostas');
// Horas das tarefas somam exatamente as horas de cada atividade
for (const [i, atv] of plano.entries()) {
  const soma = tarefas.filter((t) => t.atividadeOrdem === i).reduce((s, t) => s + parseFloat(String(t.cargaHoraria).replace(',', '.')), 0);
  assert.equal(soma, horasDe(atv), `horas das tarefas da atividade ${i + 1}`);
}
console.log('  OK: plano 34h —', plano.map((a) => `${a.cargaHoraria}`).join(' + '));

// Muitos problemas em muitas áreas → no máximo 10 atividades e NENHUM problema de fora
const AREAS = ['Finanças', 'Custos', 'Comercial', 'Estoque', 'Produção', 'Pessoas', 'Legal', 'Marketing'];
const muitos = Array.from({ length: 24 }, (_, i) => r(`P${i}`, `Possui controle ${i}?`, i % 3 === 0 ? 'Parcial' : 'Não', 1 + (i % 3), AREAS[i % AREAS.length]));
const probsMuitos = muitos.map((x, i) => prob(x.idProblema, x.area, ['Alto', 'Médio', 'Baixo'][i % 3]));
const anMuitos = analisarDiagnostico(muitos, probsMuitos, []);
assert.equal(agruparPorArea(anMuitos).length, 8);
for (const carga of [34, 40, 60, 16, 14]) {
  const p = montarAtividades(anMuitos, carga);
  verificaEstrutura(p, carga, `${carga}h/24 problemas`);
  const cobertos = new Set(p.flatMap((a) => (a.problemasTratados || []).map((t) => t.idProblema)));
  assert.equal(cobertos.size, 24, `${carga}h: todos os 24 problemas no plano`);
  const criticos = anMuitos.problemas.filter((x) => x.nivel === 'Crítico').map((x) => x.idProblema);
  assert.ok(criticos.every((id) => cobertos.has(id)), `${carga}h: todos os críticos no plano`);
  const tt = tarefasDasAtividades(p);
  assert.equal(tt.filter((t) => t.criticidade).length, 24, `${carga}h: uma tarefa por problema`);
  const somaT = tt.reduce((s, t) => s + (t.cargaHoraria ? parseFloat(t.cargaHoraria.replace(',', '.')) : 0), 0);
  assert.equal(somaT, carga, `${carga}h: horas das tarefas = carga total`);
}
console.log('  OK: 24 problemas em 8 áreas — todos no plano e nas tarefas (14, 16, 34, 40 e 60h)');

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
assert.equal(tarefasDasAtividades(conv).length, conv.length, 'plano antigo: uma tarefa por atividade');
assert.equal(conv.filter((a) => a.nome === ATIVIDADE_ANALISE_FINAL).length, 1, 'mantém uma única análise final');
assert.equal(papelDaAtividade('Relatório final e encerramento'), 'relatorio');
console.log('  OK: conversão de planos/modelos antigos');

// IA: melhora textos, mas não quebra a estrutura
const fakeAI = { models: { generateContent: async () => ({ text: JSON.stringify([
  { idProblema: 'A', acao: 'Implantar fluxo de caixa diário', solucao: 's', passos: 'passo 1; passo 2', resultadoEsperado: '100% dos lançamentos em 30 dias' },
  { idProblema: 'B', acao: 'Relatório final e encerramento', solucao: 's', passos: 'p', resultadoEsperado: 'r' },
  { idProblema: '__ferramentas__', acao: 'x', solucao: 's', passos: 'Planilha de fluxo de caixa e controle de dívidas', resultadoEsperado: '2 planilhas' },
  { idProblema: 'NAO_EXISTE', acao: 'Inventada', solucao: 's', passos: 'p', resultadoEsperado: 'r' },
]) }) } };
const ia = await enriquecerComIA(fakeAI, plano, an);
assert.equal(ia.usouIA, true);
verificaEstrutura(ia.atividades, 34, 'após IA');
assert.equal(ia.atividades[3].problemasTratados![0].acao, 'Implantar fluxo de caixa diário');
assert.match(ia.atividades[3].descricao, /passo 1; passo 2/);
assert.equal(ia.atividades[3].nome, plano[3].nome, 'nome da atividade de área não muda');
assert.equal(ia.atividades[4].nome, plano[4].nome, 'IA não pode renomear para atividade fixa');
assert.equal(ia.atividades[2].descricao, 'Planilha de fluxo de caixa e controle de dívidas');
assert.equal(ia.atividades.length, plano.length, 'IA não adiciona atividades');
const falha = await enriquecerComIA({ models: { generateContent: async () => { throw new Error('401'); } } }, plano, an);
assert.equal(falha.usouIA, false); assert.equal(falha.atividades, plano);
console.log('  OK: IA só altera textos; falha da IA mantém o plano');

// Relatório
const rel = montarTextosRelatorio(an, plano, { nomeEmpresa: 'Empresa X' });
assert.match(rel.objetivo, /^Solucionar os 4 problemas identificados no diagnóstico da empresa Empresa X \(3 críticos\), distribuídos em 4 áreas/);
console.log('   Objetivo:', rel.objetivo);
assert.match(rel.solucoesIndicadas, /PROBLEMAS IDENTIFICADOS:\nFINANÇAS \(1 crítico\):\n• \[Crítico\] Problema A\n  Ação Proposta: Implantar fluxo de caixa[\s\S]*SOLUÇÕES \/ AÇÕES PROPOSTAS:[\s\S]*PONTOS FORTES:/);
assert.match(rel.resultadosEsperados, /\[Crítico\] Problema A: Fluxo de caixa em uso/);
console.log('  OK: textos do relatório');

// Observações / contexto: "Não" nem sempre é deficiência
{
  const base = { pergunta: 'Possui controle de estoque?', problema: 'Estoque' };
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: '' }), 2, 'Não sem observação = lacuna');
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Não se aplica ao negócio (prestação de serviço)' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'O contador faz esse controle' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'MEI dispensado dessa obrigação' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Parcial', observacao: 'Já resolvido na última visita' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Cliente não controla nada e perde mercadoria' }), 2, 'observação sem indício de "ok" mantém lacuna');
  // pergunta de risco: "já possui" na observação não pode apagar um problema real
  assert.equal(grauLacuna({ pergunta: 'Possui dívidas em atraso?', problema: 'Crédito', resposta: 'Sim', observacao: 'Já possui duas dívidas em atraso' }), 2);
  // Novas expressões de observações e contexto
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Não tem necessidade no momento' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Inaplicável para o modelo da empresa' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'O próprio dono faz todo o acompanhamento diário' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Gestão familiar direta' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Vendas direto pelo WhatsApp' }), 0);
  assert.equal(grauLacuna({ ...base, resposta: 'Não', observacao: 'Trabalha apenas a vista' }), 0);
  assert.equal(grauLacuna({ pergunta: 'Possui dívidas em atraso?', problema: 'Crédito', resposta: 'Não', observacao: 'Sem dívidas, situação regular' }), 0);
  assert.equal(grauLacuna({ pergunta: 'Possui dívidas em atraso?', problema: 'Crédito', resposta: 'Sim', observacao: 'Sem dívidas, já quitou todas' }), 0);

  // Testes de classificarStatusResposta
  const st1 = classificarStatusResposta({ ...base, resposta: 'Não', observacao: 'O contador faz esse controle' });
  assert.equal(st1.tipo, 'conforme_obs');
  assert.equal(st1.isConforme, true);

  const st2 = classificarStatusResposta({ pergunta: 'O produtor comercializa apenas em feiras/festivais locais?', resposta: 'Não' });
  assert.equal(st2.tipo, 'conforme_contexto');
  assert.equal(st2.isConforme, true);

  const st3 = classificarStatusResposta({ ...base, resposta: 'Não', observacao: '' });
  assert.equal(st3.tipo, 'lacuna');
  assert.equal(st3.isConforme, false);

  // Casos reais do usuário (Comercial / Feiras / Venda institucional)
  const qFeira = {
    pergunta: 'O produtor tem interesse em participar da feira do município?',
    problema: 'Feiras e Festivais',
    resposta: 'Sim',
    observacao: 'O interesse em participar da feira municipal fortalece vendas locais, exigindo planejamento comercial e diversificação de canais.'
  };
  assert.equal(isPerguntaNegativa(qFeira.pergunta, qFeira.problema), false, 'pergunta afirmativa de feira não é negativa');
  assert.equal(grauLacuna(qFeira), 0, 'Sim com observação positiva em feira = 0');
  const stFeira = classificarStatusResposta(qFeira);
  assert.equal(stFeira.isConforme, true);
  assert.equal(stFeira.tipo, 'conforme');

  const qPnae = {
    pergunta: 'Existem parcerias com o poder público para venda institucional (PNAE, PAA)?',
    problema: 'Canais de comercialização limitados',
    resposta: 'Sim',
    observacao: 'A atuação no mercado institucional consolida um importante ponto forte de diversificação dos canais comerciais.'
  };
  assert.equal(isPerguntaNegativa(qPnae.pergunta, qPnae.problema), false, 'pergunta afirmativa de PNAE não é negativa');
  assert.equal(grauLacuna(qPnae), 0, 'Sim com observação de ponto forte em PNAE = 0');
  const stPnae = classificarStatusResposta(qPnae);
  assert.equal(stPnae.isConforme, true);
  assert.equal(stPnae.tipo, 'conforme');

  // Testes em todas as demais áreas
  // Financeiro
  const qFin = {
    pergunta: 'A empresa possui controle diário de entradas e saídas de caixa?',
    problema: 'Ausência de controle financeiro',
    resposta: 'Sim',
    observacao: 'Fluxo de caixa atualizado diariamente no sistema'
  };
  assert.equal(isPerguntaNegativa(qFin.pergunta, qFin.problema), false);
  assert.equal(grauLacuna(qFin), 0);
  assert.equal(classificarStatusResposta(qFin).isConforme, true);

  // Produção / Operações
  const qProd = {
    pergunta: 'Os processos internos estão documentados e padronizados?',
    problema: 'Falta de padronização e gargalos produtivos',
    resposta: 'Sim',
    observacao: 'Manuais de procedimentos operacionais implantados'
  };
  assert.equal(isPerguntaNegativa(qProd.pergunta, qProd.problema), false);
  assert.equal(grauLacuna(qProd), 0);
  assert.equal(classificarStatusResposta(qProd).isConforme, true);

  // Recursos Humanos
  const qRh = {
    pergunta: 'Existe um plano de treinamento e capacitação para a equipe?',
    problema: 'Ausência de capacitação',
    resposta: 'Sim',
    observacao: 'Treinamentos mensais realizados com a equipe'
  };
  assert.equal(isPerguntaNegativa(qRh.pergunta, qRh.problema), false);
  assert.equal(grauLacuna(qRh), 0);
  assert.equal(classificarStatusResposta(qRh).isConforme, true);

  // Prioridade soberana da observação:
  // Se a observação for positiva -> análise é POSITIVA (conforme / sem lacuna)
  const qNaoObsPositiva = {
    pergunta: 'A empresa possui software ERP homologado?',
    problema: 'Tecnologia',
    resposta: 'Não',
    observacao: 'Não se aplica ao porte atual, gestão simplificada em planilhas atende perfeitamente e sem erros.'
  };
  assert.equal(grauLacuna(qNaoObsPositiva), 0, 'observação positiva sobrepõe resposta Não gerando conformidade');
  const stNaoObsPos = classificarStatusResposta(qNaoObsPositiva);
  assert.equal(stNaoObsPos.isConforme, true);

  // Se a observação for negativa -> análise é NEGATIVA (lacuna a tratar), mesmo com resposta "Sim"
  const qSimObsNegativa = {
    pergunta: 'A empresa possui controle financeiro?',
    problema: 'Finanças',
    resposta: 'Sim',
    observacao: 'Apesar de afirmar que sim, não possui controle de fluxo de caixa e mistura contas pessoais com desvios frequentes.'
  };
  assert.equal(grauLacuna(qSimObsNegativa), 2, 'observação negativa sobrepõe resposta Sim gerando lacuna a tratar');
  const stSimObsNeg = classificarStatusResposta(qSimObsNegativa);
  assert.equal(stSimObsNeg.isConforme, false);
  assert.equal(stSimObsNeg.tipo, 'lacuna');

  // Veredito da IA (por id) vale para todos os consumidores e cai se a observação mudar
  const resp = r('X', 'Possui alvará sanitário?', 'Não');
  resp.observacao = 'Atividade exercida fora de estabelecimento';
  const ia = await (await import('../src/lib/domain/planoAcao')).avaliarRespostasComIA(
    { models: { generateContent: async () => ({ text: JSON.stringify([{ id: resp.id, temProblema: false }]) }) } }, [resp]);
  assert.equal(ia.get(resp.id), false);
  assert.equal(grauLacuna(resp), 0, 'veredito da IA usado sem passar o mapa');
  resp.observacao = 'Mudou o contexto';
  assert.equal(grauLacuna(resp), 2, 'veredito antigo descartado quando a observação muda');
  console.log('  OK: observações e contexto não geram lacuna indevida');
}

// Sincronização bidirecional entre Plano de Ação e Relatório
{
  const atvsIniciais: any[] = [
    { nome: '1. Diagnóstico', cargaHoraria: '4h', solucaoProposta: 'Diagnóstico', status: 'Pendente' },
    { nome: '2. Devolutiva', cargaHoraria: '2h', solucaoProposta: 'Devolutiva', status: 'Pendente' },
    {
      nome: '4. Gestão Financeira',
      cargaHoraria: '4h',
      solucaoProposta: 'Implantar fluxo de caixa',
      problemasTratados: [
        { idProblema: 'p1', problema: 'Falta de controle', solucao: 'Planilha diária', acao: 'Fluxo de caixa', passos: 'Criar planilha', resultadoEsperado: 'Caixa diário em dia', responsavel: 'Consultor' }
      ],
      status: 'Pendente'
    },
    { nome: '8. Relatório Final', cargaHoraria: '4h', solucaoProposta: 'Entrega final', status: 'Pendente' }
  ];

  // 1. Edição de tarefa do plano reflete no cronograma
  const editada = aplicarTarefaNoCronograma(atvsIniciais, {
    idProblema: 'p1',
    problema: 'Falta de controle financeiro diário',
    solucaoSugerida: 'Fluxo de caixa — Planilha automatizada',
    acoes: 'Passo A; Passo B',
    resultadoEsperado: 'Fluxo 100% conciliado',
    responsavel: 'Gestor Financeiro',
    prioridade: 'Alta'
  });
  assert.equal(editada[2].problemasTratados[0].problema, 'Falta de controle financeiro diário');
  assert.equal(editada[2].problemasTratados[0].solucao, 'Planilha automatizada');
  assert.equal(editada[2].problemasTratados[0].passos, 'Passo A; Passo B');
  assert.equal(editada[2].problemasTratados[0].resultadoEsperado, 'Fluxo 100% conciliado');
  assert.equal(editada[2].problemasTratados[0].responsavel, 'Gestor Financeiro');
  assert.equal(editada[2].prioridade, 'Alta');

  // 2. Sincronizar relatório a partir do cronograma atualizado
  const relSinc = sincronizarRelatorioComCronograma({
    objetivo: 'Realizar consultoria por meio de 4 atividades (14h) para melhoria contínua',
    solucoesIndicadas: 'PROBLEMAS IDENTIFICADOS:\n• Falta de controle\n\nSOLUÇÕES / AÇÕES PROPOSTAS:\n1. ...\n\nPONTOS FORTES:\n...',
    resultadosEsperados: 'Antigo resultado'
  }, editada);

  assert.match(relSinc.solucoesIndicadas!, /SOLUÇÕES \/ AÇÕES PROPOSTAS:[\s\S]*Gestão Financeira/);
  assert.match(relSinc.resultadosEsperados!, /Fluxo 100% conciliado/);
  assert.match(relSinc.objetivo!, /por meio de 4 atividades \(14h\)/);

  // 3. Adição de nova tarefa ao plano cria atividade antes do relatório
  const comNova = adicionarTarefaAoCronograma(editada, {
    problema: 'Nova Atividade de Vendas',
    solucaoSugerida: 'Treinamento comercial',
    acoes: 'Capacitar vendedores',
    cargaHoraria: '4h',
    resultadoEsperado: 'Vendas aumentadas',
    responsavel: 'Consultor de Vendas',
    prioridade: 'Alta'
  });
  assert.equal(comNova.length, 5);
  assert.equal(comNova[3].nome, 'Nova Atividade de Vendas');
  assert.equal(comNova[4].nome, '8. Relatório Final');

  // 4. Remoção de tarefa do plano remove do cronograma
  const comRemovida = removerTarefaDoCronograma(comNova, {
    problema: 'Nova Atividade de Vendas'
  });
  assert.equal(comRemovida.length, 4);
  assert.ok(!comRemovida.some(a => a.nome === 'Nova Atividade de Vendas'));

  console.log('  OK: sincronização bidirecional entre Plano de Ação e Relatório');
}
console.log('Todos os testes de planoAcao passaram.');

