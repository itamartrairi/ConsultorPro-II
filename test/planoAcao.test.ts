import assert from 'node:assert/strict';
import {
  analisarDiagnostico, montarAtividades, estruturarPlano, enriquecerComIA, montarTextosRelatorio,
  isPerguntaNegativa, grauLacuna, classificarStatusResposta, papelDaAtividade, horasDe, MAX_ATIVIDADES, tarefasDasAtividades, agruparPorArea,
  ATIVIDADE_DIAGNOSTICO, ATIVIDADE_DEVOLUTIVA, ATIVIDADE_FERRAMENTAS, ATIVIDADE_ANALISE_FINAL, ATIVIDADE_RELATORIO_FINAL,
  aplicarTarefaNoCronograma, removerTarefaDoCronograma, adicionarTarefaAoCronograma, sincronizarRelatorioComCronograma,
  isPerguntaFeiraOuCanalComercial, isPerguntaInteresseOuOpcao,
  isAtividadeConcluida, isAtividadeNoRelatorio,
  deduplicateTarefasPlano, findDuplicateTarefaIds,
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

  // Não em feira NÃO ESTÁ CONFORME (Lacuna Comercial a Tratar)
  const qFeiraNao = {
    pergunta: 'O produtor tem interesse em participar da feira do município?',
    problema: 'Feiras e Festivais',
    resposta: 'Não',
    observacao: ''
  };
  assert.equal(isPerguntaFeiraOuCanalComercial(qFeiraNao.pergunta), true, 'detecta feira como canal comercial');
  assert.equal(isPerguntaInteresseOuOpcao(qFeiraNao.pergunta), false, 'feira NÃO é mera opção estratégica dispensável');
  assert.equal(grauLacuna(qFeiraNao), 2, 'Não ter interesse na feira do município é lacuna a tratar');
  const stFeiraNao = classificarStatusResposta(qFeiraNao);
  assert.equal(stFeiraNao.isConforme, false, 'Não em feira do município NÃO está conforme');
  assert.equal(stFeiraNao.tipo, 'lacuna');
  assert.equal(stFeiraNao.label, 'Lacuna a Tratar (Canal Comercial / Feira)');

  // Mesmo que haja observação sugerindo canal alternativo, não ter interesse na feira do município é lacuna:
  const qFeiraNaoComObs = {
    ...qFeiraNao,
    observacao: 'Sensibilize o produtor sobre as oportunidades da feira do município para ampliar vendas e renda.'
  };
  assert.equal(grauLacuna(qFeiraNaoComObs), 2);
  assert.equal(classificarStatusResposta(qFeiraNaoComObs).isConforme, false);

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
  // Caso real da imagem: Resposta Não com observação que manda corrigir/elaborar/capacitar
  const qProtocoloHigiene = {
    pergunta: 'Existe algum protocolo escrito de higiene e segurança seguido na produção?',
    problema: 'Higiene e Segurança na Produção',
    resposta: 'Não',
    observacao: 'Elaborar procedimentos padronizados de higiene e capacitar a equipe produtiva em Boas Práticas de Fabricação.'
  };
  assert.equal(grauLacuna(qProtocoloHigiene), 2, 'Resposta Não com observação de elaborar/capacitar em boas práticas é lacuna 2');
  const stHigiene = classificarStatusResposta(qProtocoloHigiene);
  assert.equal(stHigiene.isConforme, false, 'Não deve ser conforme');
  assert.equal(stHigiene.tipo, 'lacuna', 'Tipo deve ser lacuna');
  assert.equal(stHigiene.label, 'Lacuna a Tratar (Ação Recomendada)');

  // Caso real da nova imagem: Resposta Não com observação no imperativo "Elabore um planejamento financeiro..."
  const qCapitalGiro = {
    pergunta: 'O produtor possui reserva financeira ou capital de giro para atravessar a entressafra?',
    problema: 'Gestão de Capital de Giro',
    resposta: 'Não',
    observacao: 'Elabore um planejamento financeiro focado na formação de capital de giro e reserva para a entressafra.'
  };
  assert.equal(grauLacuna(qCapitalGiro), 2, 'Resposta Não com Elabore planejamento financeiro é lacuna 2');
  const stCapital = classificarStatusResposta(qCapitalGiro);
  assert.equal(stCapital.isConforme, false, 'Não deve ser conforme');
  assert.equal(stCapital.tipo, 'lacuna');
  assert.equal(stCapital.label, 'Lacuna a Tratar (Ação Recomendada)');

  // Demais casos de observações que mandam corrigir o problema
  const qMandaCorrigir = {
    pergunta: 'A empresa possui controle de estoque?',
    problema: 'Estoque descontrolado',
    resposta: 'Não',
    observacao: 'Manda corrigir o problema implantando fichas de controle de entrada e saída.'
  };
  assert.equal(grauLacuna(qMandaCorrigir), 2);
  assert.equal(classificarStatusResposta(qMandaCorrigir).isConforme, false);

  const qNecessarioAjustar = {
    pergunta: 'Realiza precificação correta dos produtos?',
    problema: 'Margem de lucro baixa',
    resposta: 'Não',
    observacao: 'Necessário ajustar a precificação e apurar os custos fixos reais.'
  };
  assert.equal(grauLacuna(qNecessarioAjustar), 2);
  assert.equal(classificarStatusResposta(qNecessarioAjustar).isConforme, false);

  // Pergunta de risco com resposta Não, mas a observação manda corrigir
  const qRiscoComCorrecao = {
    pergunta: 'A empresa possui dívidas em atraso ou restrições?',
    problema: 'Endividamento excessivo',
    resposta: 'Não',
    observacao: 'Corrigir pendências cadastrais urgentes e renegociar débitos com fornecedores.'
  };
  assert.equal(grauLacuna(qRiscoComCorrecao), 2, 'Observação que manda corrigir se sobrepõe ao Não da pergunta de risco');
  assert.equal(classificarStatusResposta(qRiscoComCorrecao).isConforme, false);

  // Casos reais de Preceitos Ambientais e Decisões Estratégicas
  // 1. Queimadas: Resposta Não = Ponto Forte Ambiental / Sem Risco
  const qQueimadasNao = {
    pergunta: 'O produtor faz uso de queimadas no plantio ou na colheita?',
    problema: 'Impacto Ambiental no Solo',
    resposta: 'Não',
    observacao: 'A não utilização de queimadas evidencia a maturidade ambiental do produtor e a preservação do solo.'
  };
  assert.equal(grauLacuna(qQueimadasNao), 0, 'Não em queimadas é ponto forte ambiental');
  const stQueimadas = classificarStatusResposta(qQueimadasNao);
  assert.equal(stQueimadas.isConforme, true);
  assert.equal(stQueimadas.label, 'Ponto Forte (Sem Risco)');

  const qQueimadasSim = {
    pergunta: 'O produtor faz uso de queimadas no plantio ou na colheita?',
    problema: 'Impacto Ambiental no Solo',
    resposta: 'Sim',
    observacao: 'Prática de queimadas frequente no manejo'
  };
  assert.equal(grauLacuna(qQueimadasSim), 2, 'Sim em queimadas é lacuna grave');
  assert.equal(classificarStatusResposta(qQueimadasSim).isConforme, false);

  // 2. Uso de lenha/madeira: Resposta Não = Não consome biomassa florestal / Ponto Forte
  const qLenhaNao = {
    pergunta: 'Há uso de lenha/madeira na produção?',
    problema: 'Consumo de Recursos Florestais',
    resposta: 'Não',
    observacao: 'Transicione para fontes energéticas renováveis ou utilize lenha de reflorestamento certificado para reduzir impactos ambientais.'
  };
  assert.equal(grauLacuna(qLenhaNao), 0, 'Não em uso de lenha é ponto forte ambiental');
  const stLenha = classificarStatusResposta(qLenhaNao);
  assert.equal(stLenha.isConforme, true);
  assert.equal(stLenha.label, 'Ponto Forte (Sem Risco)');

  const qLenhaSim = {
    pergunta: 'Há uso de lenha/madeira na produção?',
    problema: 'Consumo de Recursos Florestais',
    resposta: 'Sim',
    observacao: 'Consome lenha nativa em caldeira'
  };
  assert.equal(grauLacuna(qLenhaSim), 2, 'Sim em uso de lenha é lacuna a tratar');

  // 3. Parceria para receber madeira de poda: Pergunta de interesse/opção voluntária
  const qPodaNao = {
    pergunta: 'O produtor tem interesse em fazer parceria com a Secretaria de Meio Ambiente para receber madeira de poda e usá-la no beneficiamento de seus produtos?',
    problema: 'Parcerias Institucionais',
    resposta: 'Não',
    observacao: 'A decisão reflete uma opção estratégica alinhada ao processo produtivo atual, sem necessidade de receber resíduos.'
  };
  assert.equal(grauLacuna(qPodaNao), 0, 'Não em pergunta de interesse opcional é decisão estratégica e não lacuna');
  const stPoda = classificarStatusResposta(qPodaNao);
  assert.equal(stPoda.isConforme, true);
  assert.equal(stPoda.label, 'Conforme (Decisão Estratégica)');

  // 4. Caso real da imagem: Licença ambiental é obrigação legal e regulatória (NÃO é mera opção estratégica!)
  const qLicencaAmbiental = {
    pergunta: 'O produtor tem interesse em tirar licença ambiental?',
    problema: 'Licenciamento Ambiental',
    resposta: 'Não',
    observacao: 'Oriente o produtor sobre a importância do licenciamento ambiental para a regularização legal e manejo sustentável.'
  };
  assert.equal(grauLacuna(qLicencaAmbiental), 2, 'Não em licença ambiental é lacuna a tratar 2');
  const stLicenca = classificarStatusResposta(qLicencaAmbiental);
  assert.equal(stLicenca.isConforme, false, 'Não deve ser conforme');
  assert.equal(stLicenca.tipo, 'lacuna');
  assert.equal(stLicenca.label, 'Lacuna a Tratar (Exigência Legal/Regulatória)');

  // 5. Testes em todas as demais áreas com obrigações legais / regulatórias:
  // Alvará de funcionamento
  const qAlvaraFunc = {
    pergunta: 'A empresa possui alvará de funcionamento regularizado?',
    problema: 'Formalização e Regularidade',
    resposta: 'Não',
    observacao: ''
  };
  assert.equal(grauLacuna(qAlvaraFunc), 2);
  assert.equal(classificarStatusResposta(qAlvaraFunc).isConforme, false);
  assert.equal(classificarStatusResposta(qAlvaraFunc).label, 'Lacuna a Tratar (Exigência Legal/Regulatória)');

  // Vigilância sanitária
  const qAlvaraSanit = {
    pergunta: 'A empresa possui alvará sanitário da Vigilância Sanitária?',
    problema: 'Adequação Sanitária',
    resposta: 'Não',
    observacao: 'Providenciar adequações físicas e solicitar vistoria da vigilância sanitária.'
  };
  assert.equal(grauLacuna(qAlvaraSanit), 2);
  assert.equal(classificarStatusResposta(qAlvaraSanit).isConforme, false);

  // Fiscal / Tributário (Emissão de nota fiscal)
  const qNotaFiscal = {
    pergunta: 'A empresa emite nota fiscal em 100% das vendas realizadas?',
    problema: 'Gestão Fiscal',
    resposta: 'Não',
    observacao: 'Regularize a emissão de nota fiscal para evitar autuações da Receita.'
  };
  assert.equal(grauLacuna(qNotaFiscal), 2);
  assert.equal(classificarStatusResposta(qNotaFiscal).isConforme, false);

  // Trabalhista e Segurança (EPIs / Registro em carteira)
  const qEpi = {
    pergunta: 'Todos os funcionários utilizam os EPIs obrigatórios?',
    problema: 'Segurança do Trabalho',
    resposta: 'Não',
    observacao: 'Forneça os EPIs e treine a equipe conforme as normas regulamentadoras NR-6.'
  };
  assert.equal(grauLacuna(qEpi), 2);
  assert.equal(classificarStatusResposta(qEpi).isConforme, false);

  // Veredito da IA (por id) vale para todos os consumidores e cai se a observação mudar
  const resp = r('X', 'Possui alvará sanitário?', 'Não');
  resp.observacao = 'Atividade exercida fora de estabelecimento';
  const ia = await (await import('../src/lib/domain/planoAcao')).avaliarRespostasComIA(
    { models: { generateContent: async () => ({ text: JSON.stringify([{ id: resp.id, temProblema: false }]) }) } }, [resp]);
  assert.equal(ia.get(resp.id), false);
  assert.equal(grauLacuna(resp), 0, 'veredito da IA usado sem passar o mapa');
  // Casos reais de conformidade e maturidade com resposta Sim (ex.: qualidade da cajuína, padronização, controle operacional)
  const qCajuinaQualidade = {
    pergunta: 'Existe uma análise definida para garantir a qualidade da cajuína em cada lote?',
    problema: 'Controle de Qualidade',
    resposta: 'Sim',
    observacao: 'A análise de qualidade em cada lote demonstra maturidade operacional e assegura a padronização da cajuína.'
  };
  assert.equal(grauLacuna(qCajuinaQualidade), 0, 'Sim em controle de qualidade da cajuína é ponto forte');
  assert.equal(classificarStatusResposta(qCajuinaQualidade).isConforme, true);
  assert.equal(classificarStatusResposta(qCajuinaQualidade).label, 'Ponto Forte (Conforme)');

  const qCajuinaParametros = {
    pergunta: 'Existe controle de tempo, temperatura e proporção de açúcar em cada lote produzido?',
    problema: 'Parâmetros Produtivos',
    resposta: 'Sim',
    observacao: 'O controle de parâmetros produtivos evidencia excelente maturidade operacional e garante a padronização da cajuína.'
  };
  assert.equal(grauLacuna(qCajuinaParametros), 0, 'Sim em controle de parâmetros da cajuína é ponto forte');
  assert.equal(classificarStatusResposta(qCajuinaParametros).isConforme, true);
  assert.equal(classificarStatusResposta(qCajuinaParametros).label, 'Ponto Forte (Conforme)');

  // Outros casos da biblioteca onde resposta Sim é acompanhada de observação afirmativa de boas práticas
  const qBoasPraticasEstoque = {
    pergunta: 'A empresa possui controle de estoque?',
    problema: 'Gestão de Estoques',
    resposta: 'Sim',
    observacao: 'O controle de estoque assegura a organização dos insumos e otimiza a rotatividade.'
  };
  assert.equal(grauLacuna(qBoasPraticasEstoque), 0);
  assert.equal(classificarStatusResposta(qBoasPraticasEstoque).isConforme, true);

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

// 12. Regra de seleção "No Relatório" exclusiva para status Concluído
{
  // Concluído é elegível por padrão (true)
  assert.equal(isAtividadeConcluida('Concluído'), true);
  assert.equal(isAtividadeConcluida('CONCLUÍDO'), true);
  assert.equal(isAtividadeNoRelatorio({ status: 'Concluído' }), true);
  assert.equal(isAtividadeNoRelatorio({ status: 'Concluído', incluirNoRelatorio: true }), true);
  // Concluído desmarcado explicitamente pelo usuário
  assert.equal(isAtividadeNoRelatorio({ status: 'Concluído', incluirNoRelatorio: false }), false);

  // Demais status NUNCA são permitidos no relatório, mesmo se incluirNoRelatorio vier true
  const outrosStatus = ['Pendente', 'Em Andamento', 'Atrasado', 'Não Iniciado', 'Cancelado', undefined];
  for (const st of outrosStatus) {
    assert.equal(isAtividadeConcluida(st), false, `Status ${st} não deve ser considerado concluído`);
    assert.equal(isAtividadeNoRelatorio({ status: st as any }), false, `Status ${st} não deve constar no relatório`);
    assert.equal(isAtividadeNoRelatorio({ status: st as any, incluirNoRelatorio: true }), false, `Status ${st} com incluirNoRelatorio:true deve ser bloqueado`);
  }

  // Ao sincronizar tarefa com mudança de status via aplicarTarefaNoCronograma
  const atvs = [
    { nome: 'Atividade 1', status: 'Pendente', cargaHoraria: '4h', solucaoProposta: 'Sol 1', incluirNoRelatorio: false }
  ];
  const atualizadaConcluida = aplicarTarefaNoCronograma(atvs as any, {
    problema: 'Atividade 1',
    status: 'Concluído'
  });
  assert.equal(atualizadaConcluida[0].status, 'Concluído');
  assert.equal(atualizadaConcluida[0].incluirNoRelatorio, true);
  assert.equal(isAtividadeNoRelatorio(atualizadaConcluida[0]), true);

  const revertidaPendente = aplicarTarefaNoCronograma(atualizadaConcluida, {
    problema: 'Atividade 1',
    status: 'Em Andamento'
  });
  assert.equal(revertidaPendente[0].status, 'Em Andamento');
  assert.equal(revertidaPendente[0].incluirNoRelatorio, false);
  assert.equal(isAtividadeNoRelatorio(revertidaPendente[0]), false);

  console.log('  OK: seleção "No Relatório" exclusiva para status Concluído e bloqueio dos demais status');
}

// Teste de deduplicação de tarefas do Plano de Ação
{
  const tarefasDuplicadas: any[] = [
    {
      id: 'doc_firestore_1',
      diagnosticoId: 'diag_cajuina_1',
      atividadeOrdem: 0,
      atividade: 'Diagnóstico Empresarial',
      problema: 'Diagnóstico Empresarial',
      acoes: 'Levantamento de lacunas',
      status: 'Concluído',
      cargaHoraria: '4h',
      responsavel: 'Vania Maria',
      ordem: 0
    },
    {
      id: 'tarefa_diag_cajuina_1_0',
      diagnosticoId: 'diag_cajuina_1',
      atividadeOrdem: 0,
      atividade: 'Diagnóstico Empresarial',
      problema: 'Diagnóstico Empresarial',
      acoes: 'Levantamento de lacunas',
      status: 'Pendente',
      cargaHoraria: '4h',
      responsavel: 'Consultor',
      ordem: 0
    },
    {
      id: 'doc_firestore_2',
      diagnosticoId: 'diag_cajuina_1',
      atividadeOrdem: 1,
      atividade: 'Devolutiva do Diagnóstico',
      problema: 'Devolutiva do Diagnóstico',
      status: 'Concluído',
      cargaHoraria: '3h',
      ordem: 1
    },
    {
      id: 'tarefa_diag_cajuina_1_1',
      diagnosticoId: 'diag_cajuina_1',
      atividadeOrdem: 1,
      atividade: 'Devolutiva do Diagnóstico',
      problema: 'Devolutiva do Diagnóstico',
      status: 'Pendente',
      cargaHoraria: '3h',
      ordem: 1
    },
    {
      id: 'doc_firestore_3_a',
      diagnosticoId: 'diag_cajuina_1',
      idProblema: 'PROD_1',
      atividadeOrdem: 2,
      atividade: 'Produção: Resolução',
      problema: 'Falta de padronização no processo de fabricação da cajuína',
      status: 'Em Andamento',
      evidenciaUrl: 'data:image/png;base64,xxx',
      ordem: 2
    },
    {
      id: 'doc_firestore_3_b',
      diagnosticoId: 'diag_cajuina_1',
      idProblema: 'PROD_1',
      atividadeOrdem: 2,
      atividade: 'Produção: Resolução',
      problema: 'Falta de padronização no processo de fabricação da cajuína',
      status: 'Pendente',
      ordem: 2
    }
  ];

  const limpas = deduplicateTarefasPlano(tarefasDuplicadas);
  assert.equal(limpas.length, 3, 'Deve reduzir 6 tarefas (3 pares duplicados) para exatamente 3 tarefas únicas');
  assert.equal(limpas[0].id, 'doc_firestore_1', 'Deve preservar o ID definitivo da nuvem');
  assert.equal(limpas[0].status, 'Concluído', 'Status Concluído deve prevalecer sobre Pendente');
  assert.equal(limpas[0].responsavel, 'Vania Maria', 'Responsável preenchido deve prevalecer');
  assert.equal(limpas[1].id, 'doc_firestore_2');
  assert.equal(limpas[2].status, 'Em Andamento');
  assert.equal(limpas[2].evidenciaUrl, 'data:image/png;base64,xxx', 'Evidência deve ser preservada na mesclagem');

  const dupIds = findDuplicateTarefaIds(tarefasDuplicadas);
  assert.deepEqual(dupIds, ['tarefa_diag_cajuina_1_0', 'tarefa_diag_cajuina_1_1', 'doc_firestore_3_b'], 'Deve identificar com precisão os IDs redundantes');

  console.log('  OK: deduplicação de tarefas do plano de ação e identificação de IDs redundantes');
}

console.log('Todos os testes de planoAcao passaram.');


