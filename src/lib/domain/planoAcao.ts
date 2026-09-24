/**
 * Geração do PLANO DE AÇÃO e dos textos do RELATÓRIO a partir do diagnóstico.
 *
 * Regras:
 *  1. Os problemas são classificados pela criticidade real das respostas (lacunas × peso)
 *     combinada com o impacto cadastrado do problema (Alto / Médio / Baixo).
 *  2. O plano resolve primeiro os problemas CRÍTICOS, depois os de prioridade alta e,
 *     se sobrar carga horária, os moderados (agrupados numa atividade complementar).
 *  3. A PRIMEIRA atividade é sempre o diagnóstico e a ÚLTIMA é sempre o preparo dos
 *     relatórios e a apresentação dos resultados — independentemente da IA.
 *  4. A IA (opcional) só reescreve os textos de cada atividade; a estrutura é garantida
 *     aqui, então o plano sai correto mesmo sem IA ou se ela falhar.
 */
import type { AtividadeCronograma, Problema, Resposta, Solucao } from '../../types/domain';
import { Type } from '../ai/schemaTypes';

export type NivelCriticidade = 'Crítico' | 'Alto' | 'Moderado';

export interface Lacuna {
  pergunta: string;
  resposta: string;
  observacao?: string;
}

export interface ProblemaAnalisado {
  idProblema: string;
  problema: string;
  area: string;
  impacto: string;
  nivel: NivelCriticidade;
  /** 0–100: quanto maior, mais urgente. */
  pontuacao: number;
  /** Fração das respostas que indicam o problema (0–1), ponderada pelo peso. */
  intensidade: number;
  lacunas: Lacuna[];
  solucao?: Solucao;
}

export interface AnaliseDiagnostico {
  totalRespondidas: number;
  problemas: ProblemaAnalisado[]; // só os que têm lacuna, do mais ao menos crítico
  pontosFortes: { problema: string; area: string }[];
  scoreGeral: number; // 0–100 (100 = sem lacunas)
}

// ---------------------------------------------------------------------------------------
// Estrutura fixa do plano
// ---------------------------------------------------------------------------------------
export const MAX_ATIVIDADES = 10;
export const ATIVIDADE_DIAGNOSTICO = 'Diagnóstico empresarial';
export const ATIVIDADE_DEVOLUTIVA = 'Devolutiva do diagnóstico';
export const ATIVIDADE_FERRAMENTAS = 'Desenvolvimento de ferramentas de gestão (planilhas e aplicativos)';
export const ATIVIDADE_ANALISE_FINAL = 'Análise final de atividades da consultoria';
export const ATIVIDADE_RELATORIO_FINAL = 'Preparo de relatórios e apresentação dos resultados';
/** Atividades fixas: 3 no início + 2 no fim. Sobram até 5 para os problemas críticos. */
export const MAX_ATIVIDADES_PROBLEMA = MAX_ATIVIDADES - 5;

export type PapelAtividade = 'diagnostico' | 'devolutiva' | 'ferramentas' | 'problema' | 'analise' | 'relatorio';

const norm = (s: unknown) =>
  String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();

/** Identifica o papel de uma atividade pelo nome (aceita os nomes antigos do app). */
export function papelDaAtividade(nome?: string): PapelAtividade {
  const n = norm(nome);
  if (n.includes('devolutiva')) return 'devolutiva';
  if (n.includes('analise final')) return 'analise';
  if (
    n.includes('relatorio final') ||
    n.includes('preparo de relatorio') ||
    n.includes('apresentacao dos resultados') ||
    n.includes('encerramento')
  ) return 'relatorio';
  if (
    n.includes('ferramentas de gestao') ||
    n.includes('desenvolvimento de sistema') ||
    n.includes('desenvolvimento de ferramenta') ||
    (n.includes('desenvolvimento') && (n.includes('planilha') || n.includes('aplicativo')))
  ) return 'ferramentas';
  if (
    n === 'diagnostico' ||
    n.startsWith('diagnostico empresarial') ||
    n.includes('diagnostico inicial') ||
    n.includes('entendimento da demanda')
  ) return 'diagnostico';
  return 'problema';
}

export const isAtividadeDiagnostico = (nome?: string) => papelDaAtividade(nome) === 'diagnostico';
export const isAtividadeRelatorioFinal = (nome?: string) => papelDaAtividade(nome) === 'relatorio';


// ---------------------------------------------------------------------------------------
// Semântica das perguntas
// ---------------------------------------------------------------------------------------
/**
 * Pergunta que investiga a OCORRÊNCIA de um problema ("Já foi rejeitado…?", "Possui
 * dívidas…?"). Nelas, "Sim" é o problema. Nas demais (controles e boas práticas), o
 * problema é o "Não".
 */
const PADRAO_PERGUNTA_NEGATIVA = new RegExp(
  [
    '\\bja (sofreu|teve|foi|tentou|perdeu|atrasou)',
    '\\bfoi (rejeitad|negad|recusad|autuad|multad|notificad|embargad)',
    '\\b(sofre|enfrenta|apresenta) (com |)(perda|prejuizo|atraso|dificuldade|problema|queda|falta|reclama|mortalidade|inadimpl)',
    '\\b(ha|existe|existem|ocorre|ocorrem) (atraso|perda|desperdicio|reclama|inadimpl|divida|pendencia|falta|mortalidade|retrabalho|devolu|acidente)',
    '\\b(possui|tem) (dividas?|pendencias?|inadimplencia|restricao|restricoes|multas?|dificuldade|problemas?|nome negativado)',
    '\\besta (endividad|inadimplent|negativad|irregular)',
    '\\bdepende (apenas|somente|exclusivamente)',
  ].join('|')
);

export function isPerguntaNegativa(pergunta?: string): boolean {
  return PADRAO_PERGUNTA_NEGATIVA.test(norm(pergunta));
}

/** 2 = lacuna total, 1 = parcial, 0 = sem lacuna, null = não respondida. */
export function grauLacuna(r: Pick<Resposta, 'resposta' | 'pergunta'>): number | null {
  const v = r.resposta;
  if (v !== 'Sim' && v !== 'Não' && v !== 'Parcial') return null;
  if (v === 'Parcial') return 1;
  const negativa = isPerguntaNegativa(r.pergunta);
  if (negativa) return v === 'Sim' ? 2 : 0;
  return v === 'Não' ? 2 : 0;
}

function pesoImpacto(impacto?: string): number {
  const n = norm(impacto);
  if (n.startsWith('alt') || n.startsWith('critic')) return 3;
  if (n.startsWith('baix')) return 1;
  return 2; // Médio ou não informado
}

const ORDEM_NIVEL: Record<NivelCriticidade, number> = { Crítico: 3, Alto: 2, Moderado: 1 };

// ---------------------------------------------------------------------------------------
// Análise
// ---------------------------------------------------------------------------------------
export function analisarDiagnostico(
  respostas: Resposta[],
  problemas: Problema[] = [],
  solucoes: Solucao[] = []
): AnaliseDiagnostico {
  const grupos = new Map<string, Resposta[]>();
  let totalRespondidas = 0;
  let somaLacuna = 0;
  let somaMax = 0;

  for (const r of respostas || []) {
    if (!r) continue;
    const g = grauLacuna(r);
    if (g === null) continue;
    totalRespondidas++;
    const peso = Number(r.peso) > 0 ? Number(r.peso) : 1;
    somaLacuna += g * peso;
    somaMax += 2 * peso;
    const chave = r.idProblema || norm(r.problema) || 'sem_problema';
    if (!grupos.has(chave)) grupos.set(chave, []);
    grupos.get(chave)!.push(r);
  }

  const analisados: ProblemaAnalisado[] = [];
  const pontosFortes: AnaliseDiagnostico['pontosFortes'] = [];

  for (const [chave, lista] of grupos) {
    const ref = lista[0];
    const cadastro =
      problemas.find((p) => p.id === ref.idProblema) ||
      problemas.find((p) => norm(p.descricao_problemas) === norm(ref.problema));
    const nome = cadastro?.descricao_problemas || ref.problema || 'Problema identificado';
    const area = cadastro?.area || ref.area || 'Geral';
    const impacto = cadastro?.impacto || 'Médio';

    let lac = 0;
    let max = 0;
    const lacunas: Lacuna[] = [];
    for (const r of lista) {
      const g = grauLacuna(r)!;
      const peso = Number(r.peso) > 0 ? Number(r.peso) : 1;
      lac += g * peso;
      max += 2 * peso;
      if (g > 0) lacunas.push({ pergunta: r.pergunta, resposta: r.resposta || '', observacao: r.observacao || undefined });
    }
    const intensidade = max > 0 ? lac / max : 0;

    if (intensidade === 0) {
      pontosFortes.push({ problema: nome, area });
      continue;
    }

    const imp = pesoImpacto(impacto);
    const temLacunaTotal = lista.some((r) => grauLacuna(r) === 2);
    let nivel: NivelCriticidade = 'Moderado';
    if ((imp >= 3 && temLacunaTotal) || (intensidade >= 0.75 && imp >= 2)) nivel = 'Crítico';
    else if ((intensidade >= 0.5 && imp >= 2) || imp >= 3) nivel = 'Alto';

    const solucao =
      solucoes.find((s) => s.idProblema && s.idProblema === (cadastro?.id || ref.idProblema)) ||
      solucoes.find((s) => norm(s.problema) === norm(nome));

    analisados.push({
      idProblema: cadastro?.id || ref.idProblema || chave,
      problema: nome,
      area,
      impacto,
      nivel,
      pontuacao: Math.round(intensidade * 100 * (0.5 + imp / 6)),
      intensidade,
      lacunas,
      solucao,
    });
  }

  analisados.sort(
    (a, b) => ORDEM_NIVEL[b.nivel] - ORDEM_NIVEL[a.nivel] || b.pontuacao - a.pontuacao || a.problema.localeCompare(b.problema)
  );

  return {
    totalRespondidas,
    problemas: analisados,
    pontosFortes,
    scoreGeral: somaMax > 0 ? Math.round((1 - somaLacuna / somaMax) * 100) : 0,
  };
}

// ---------------------------------------------------------------------------------------
// Plano de ação (atividades do cronograma)
// ---------------------------------------------------------------------------------------
const PRIORIDADE: Record<NivelCriticidade, AtividadeCronograma['prioridade']> = {
  Crítico: 'Alta',
  Alto: 'Alta',
  Moderado: 'Média',
};

/** Ferramentas de gestão sugeridas conforme a área dos problemas. */
const FERRAMENTAS_POR_AREA: [RegExp, string][] = [
  [/financ|caixa|contab|tribut|fiscal|credito|divida|inadimpl/, 'planilha de fluxo de caixa e DRE gerencial'],
  [/custo|preco|precifica|margem/, 'planilha de custos e formação de preço de venda'],
  [/comerc|venda|cliente|mercado|marketing|atendimento/, 'controle de vendas, clientes e metas comerciais'],
  [/estoque|compra|suprimento|insumo|almoxarif/, 'controle de estoque e compras'],
  [/produc|operac|process|qualidade|manejo|logist/, 'controle de produção e indicadores operacionais'],
  [/pessoa|rh|equipe|gestao de pessoas|colaborador/, 'controle de rotinas, escalas e desempenho da equipe'],
  [/ambient|legal|juridic|licen|document|regulariz|sanitar/, 'checklist de conformidade e agenda de vencimentos'],
  [/estrateg|planejamento|gestao|governanca|inovac/, 'painel de indicadores e plano de metas'],
];

export function ferramentasSugeridas(analise?: AnaliseDiagnostico): string[] {
  const out: string[] = [];
  for (const p of analise?.problemas || []) {
    const alvo = norm(`${p.area} ${p.problema}`);
    for (const [re, ferramenta] of FERRAMENTAS_POR_AREA) {
      if (re.test(alvo) && !out.includes(ferramenta)) out.push(ferramenta);
    }
  }
  if (out.length === 0) out.push('planilha de fluxo de caixa e DRE gerencial', 'painel de indicadores e plano de metas');
  return out.slice(0, 4);
}

function listaTexto(itens: string[]): string {
  if (itens.length <= 1) return itens[0] || '';
  return `${itens.slice(0, -1).join(', ')} e ${itens[itens.length - 1]}`;
}

export function atividadeDiagnostico(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const areas = Array.from(new Set((analise?.problemas || []).map((p) => p.area))).slice(0, 5);
  return {
    nome: ATIVIDADE_DIAGNOSTICO,
    descricao:
      'Entendimento da demanda, aplicação do questionário de diagnóstico, levantamento de dados e documentos e ' +
      `análise da situação atual da empresa${areas.length ? ` (áreas: ${areas.join(', ')})` : ''}.`,
    cargaHoraria: '4h',
    solucaoProposta: 'Diagnóstico empresarial com levantamento das lacunas de gestão',
    resultadoEsperado: 'Diagnóstico aplicado, com 100% das perguntas respondidas e problemas mapeados por área.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeDevolutiva(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const criticos = (analise?.problemas || []).filter((p) => p.nivel === 'Crítico').map((p) => p.problema);
  const prioritarios = criticos.length ? criticos : (analise?.problemas || []).slice(0, 3).map((p) => p.problema);
  return {
    nome: ATIVIDADE_DEVOLUTIVA,
    descricao:
      'Apresentação dos resultados do diagnóstico ao empresário, validação dos problemas identificados, ' +
      'priorização conjunta e pactuação do plano de ação' +
      (prioritarios.length ? `, com foco em: ${prioritarios.slice(0, 5).join('; ')}.` : '.'),
    cargaHoraria: '2h',
    solucaoProposta: 'Devolutiva do diagnóstico e pactuação do plano de ação com o cliente',
    resultadoEsperado: 'Problemas críticos validados pelo empresário e plano de ação aprovado.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeFerramentas(analise?: AnaliseDiagnostico): AtividadeCronograma {
  const ferramentas = ferramentasSugeridas(analise);
  return {
    nome: ATIVIDADE_FERRAMENTAS,
    descricao:
      `Desenvolvimento e implantação de ferramentas de gestão sob medida: ${listaTexto(ferramentas)}. ` +
      'Inclui parametrização com os dados da empresa e treinamento do responsável pelo uso.',
    cargaHoraria: '8h',
    solucaoProposta: 'Planilhas e aplicativos de gestão personalizados para os problemas identificados',
    resultadoEsperado: `Ferramentas implantadas e em uso na rotina da empresa (${ferramentas.length} controle(s) ativo(s)).`,
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Alta',
  };
}

export function atividadeAnaliseFinal(): AtividadeCronograma {
  return {
    nome: ATIVIDADE_ANALISE_FINAL,
    descricao:
      'Análise crítica de todas as atividades executadas: verificação das entregas, mensuração dos indicadores ' +
      'em relação às metas do plano e registro das recomendações para a continuidade após a consultoria.',
    cargaHoraria: '2h',
    solucaoProposta: 'Análise final e consolidação dos resultados da consultoria',
    resultadoEsperado: 'Checklist de entregas concluído e indicadores de cada atividade apurados em relação às metas.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

export function atividadeRelatorioFinal(): AtividadeCronograma {
  return {
    nome: ATIVIDADE_RELATORIO_FINAL,
    descricao:
      'Consolidação das atividades executadas e dos indicadores alcançados, elaboração do relatório técnico da ' +
      'consultoria e apresentação dos resultados e recomendações ao empresário.',
    cargaHoraria: '4h',
    solucaoProposta: 'Elaboração dos relatórios e apresentação dos resultados',
    resultadoEsperado: 'Relatório final entregue e resultados apresentados e validados com o cliente.',
    responsavel: 'Consultor',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

function atividadeDoProblema(p: ProblemaAnalisado): AtividadeCronograma {
  const s = p.solucao;
  const lacunasTxt = p.lacunas
    .slice(0, 3)
    .map((l) => `"${l.pergunta}" → ${l.resposta}${l.observacao ? ` (${l.observacao})` : ''}`)
    .join('; ');
  return {
    nome: s?.solucao_recomendada ? `${p.problema}: ${s.solucao_recomendada}` : `Solução para: ${p.problema}`,
    descricao:
      s?.acoes_sugeridas ||
      `Análise das causas, definição e implantação, junto com o empresário, das correções para "${p.problema}" (${p.area}). ` +
        `Pontos do diagnóstico: ${lacunasTxt}.`,
    cargaHoraria: '4h',
    solucaoProposta: s?.solucao_recomendada || `Plano de correção para ${p.problema.toLowerCase()}`,
    resultadoEsperado:
      [s?.resultado_esperado, s?.kpis_sugeridos ? `Indicadores: ${s.kpis_sugeridos}` : ''].filter(Boolean).join(' — ') ||
      `Problema "${p.problema}" solucionado, com controle implantado e acompanhado mensalmente.`,
    responsavel: s?.responsavel_sugerido || 'Consultor/Cliente',
    idProblema: p.idProblema,
    status: 'Pendente',
    prioridade: PRIORIDADE[p.nivel],
    area: p.area,
    criticidade: p.nivel,
  };
}

function atividadeComplementar(extras: { nome: string; nivel?: string }[]): AtividadeCronograma {
  return {
    nome: 'Ações complementares para os demais problemas identificados',
    descricao: `Orientação e plano de correção para: ${extras.map((p) => (p.nivel ? `${p.nome} (${p.nivel.toLowerCase()})` : p.nome)).join('; ')}.`,
    cargaHoraria: '4h',
    solucaoProposta: 'Orientações técnicas e plano de correção para os problemas complementares',
    resultadoEsperado: 'Empresário orientado e com responsável e prazo definidos para cada problema complementar.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: extras.some((p) => p.nivel && p.nivel !== 'Moderado') ? 'Alta' : 'Média',
  };
}

function atividadeSemLacunas(): AtividadeCronograma {
  return {
    nome: 'Fortalecimento dos controles e indicadores de gestão',
    descricao: 'O diagnóstico não apontou lacunas relevantes: revisão dos controles existentes, definição de metas e rotina de acompanhamento de indicadores.',
    cargaHoraria: '4h',
    solucaoProposta: 'Rotina de acompanhamento de indicadores e metas',
    resultadoEsperado: 'Metas definidas e indicadores acompanhados mensalmente pelo empresário.',
    responsavel: 'Consultor/Cliente',
    status: 'Pendente',
    prioridade: 'Média',
  };
}

/** Horas (número) de uma atividade. */
export function horasDe(a?: Pick<AtividadeCronograma, 'cargaHoraria'>): number {
  const n = parseInt(String(a?.cargaHoraria || '').replace(/\D/g, ''), 10);
  return Number.isFinite(n) ? n : 0;
}

export function parseCargaHoraria(total?: string | number, padrao = 34): number {
  const n = typeof total === 'number' ? total : parseInt(String(total || '').replace(/\D/g, ''), 10);
  return Number.isFinite(n) && n > 0 ? n : padrao;
}

/** Quantas atividades de problema cabem na carga horária (mínimo 2h cada, máximo 6). */
export function capacidadeAtividades(cargaHorariaTotal: number): number {
  const minimoFixo = 12; // diagnóstico 2h + devolutiva 2h + ferramentas 4h + análise final 2h + relatório 2h
  return Math.max(1, Math.min(MAX_ATIVIDADES_PROBLEMA, Math.floor((cargaHorariaTotal - minimoFixo) / 2)));
}

/**
 * Distribui a carga horária total entre as atividades, somando EXATAMENTE o total.
 * Padrão: diagnóstico 4h, devolutiva 2h, ferramentas 8h, análise final 2h (fixa), relatório 4h
 * e 4h por problema. Com pouca carga, reduz até os mínimos (2h, 2h, 4h, 2h, 2h e 2h por problema).
 */
export function distribuirCargaHoraria(atividades: AtividadeCronograma[], cargaHorariaTotal: number): AtividadeCronograma[] {
  const lista = atividades.map((a) => ({ ...a }));
  const papeis = lista.map((a) => papelDaAtividade(a.nome));
  const idx = (p: PapelAtividade) => papeis.indexOf(p);
  const probs = papeis.map((p, i) => (p === 'problema' ? i : -1)).filter((i) => i >= 0);
  const n = probs.length;

  const padrao: Record<string, number> = { diagnostico: 4, devolutiva: 2, ferramentas: 8, analise: 2, relatorio: 4 };
  const minimo: Record<string, number> = { diagnostico: 2, devolutiva: 2, ferramentas: 4, analise: 2, relatorio: 2 };
  const fixos = (['diagnostico', 'devolutiva', 'ferramentas', 'analise', 'relatorio'] as const).filter((p) => idx(p) >= 0);
  const horas = new Map<number, number>();

  const somaPadrao = fixos.reduce((s, p) => s + padrao[p], 0);
  const somaMin = fixos.reduce((s, p) => s + minimo[p], 0);
  let total = Math.max(cargaHorariaTotal, somaMin + 2 * n);

  if (total >= somaPadrao + 2 * n) {
    fixos.forEach((p) => horas.set(idx(p), padrao[p]));
    let resto = total - somaPadrao;
    if (n > 0) {
      const base = Math.min(4, Math.floor(resto / n));
      probs.forEach((i) => horas.set(i, base));
      resto -= base * n;
      for (let k = 0; resto > 0 && k < n; k++) if (horas.get(probs[k])! < 4) { horas.set(probs[k], horas.get(probs[k])! + 1); resto--; }
    }
    // Sobra de horas: ferramentas (até 12h), depois problemas (até 8h), depois diagnóstico e relatório.
    const alvos: [number, number][] = [
      [idx('ferramentas'), 12],
      ...probs.map((i) => [i, 8] as [number, number]),
      [idx('diagnostico'), 8],
      [idx('relatorio'), 8],
      [idx('devolutiva'), 4],
      // A análise final fica sempre com 2h.
    ].filter(([i]) => i >= 0) as [number, number][];
    let guarda = 0;
    while (resto > 0 && guarda++ < 1000) {
      let aplicou = false;
      for (const [i, max] of alvos) {
        if (resto <= 0) break;
        if ((horas.get(i) || 0) < max) { horas.set(i, (horas.get(i) || 0) + 1); resto--; aplicou = true; }
      }
      if (!aplicou) { const alvo = idx('ferramentas') >= 0 ? idx('ferramentas') : probs[0] ?? 0; horas.set(alvo, (horas.get(alvo) || 0) + resto); resto = 0; }
    }
  } else {
    probs.forEach((i) => horas.set(i, 2));
    fixos.forEach((p) => horas.set(idx(p), minimo[p]));
    let resto = total - 2 * n - somaMin;
    for (const p of ['ferramentas', 'diagnostico', 'relatorio', 'devolutiva'] as const) {
      const i = idx(p);
      if (i < 0) continue;
      const add = Math.min(resto, padrao[p] - minimo[p]);
      horas.set(i, horas.get(i)! + add);
      resto -= add;
    }
  }

  lista.forEach((a, i) => { if (horas.has(i)) a.cargaHoraria = `${horas.get(i)}h`; });
  return lista;
}

/**
 * Normaliza QUALQUER lista (gerada, replicada, de modelo ou editada) para a estrutura:
 *   1. Diagnóstico · 2. Devolutiva · 3. Ferramentas de gestão · 4..8 problemas ·
 *   penúltima: Análise final (2h) · última: Relatórios.
 * Mantém no máximo 10 atividades (excedentes viram "Ações complementares") e ajusta as horas.
 */
export function estruturarPlano(
  lista: AtividadeCronograma[],
  cargaHorariaTotal: number | string | undefined,
  analise?: AnaliseDiagnostico
): AtividadeCronograma[] {
  const total = parseCargaHoraria(cargaHorariaTotal);
  const itens = (lista || []).filter((a) => a && (a.nome || a.solucaoProposta));
  const pega = (p: PapelAtividade) => itens.find((a) => papelDaAtividade(a.nome) === p);

  const diag = pega('diagnostico');
  const devol = pega('devolutiva');
  const ferr = pega('ferramentas');
  const analiseFinal = pega('analise');
  const rel = pega('relatorio');
  let meio = itens.filter((a) => papelDaAtividade(a.nome) === 'problema');

  const limite = Math.min(MAX_ATIVIDADES_PROBLEMA, capacidadeAtividades(total));
  if (meio.length > limite) {
    const extras = meio.slice(limite - 1);
    meio = [...meio.slice(0, limite - 1), atividadeComplementar(extras.map((e) => ({ nome: e.nome, nivel: (e as any).criticidade })))];
  }
  if (meio.length === 0) meio = [atividadeSemLacunas()];

  const estruturado: AtividadeCronograma[] = [
    diag ? { ...diag, nome: ATIVIDADE_DIAGNOSTICO } : atividadeDiagnostico(analise),
    devol ? { ...devol, nome: ATIVIDADE_DEVOLUTIVA } : atividadeDevolutiva(analise),
    ferr ? { ...ferr, nome: ATIVIDADE_FERRAMENTAS } : atividadeFerramentas(analise),
    ...meio,
    analiseFinal ? { ...analiseFinal, nome: ATIVIDADE_ANALISE_FINAL } : atividadeAnaliseFinal(),
    rel ? { ...rel, nome: ATIVIDADE_RELATORIO_FINAL } : atividadeRelatorioFinal(),
  ];
  return distribuirCargaHoraria(estruturado, total).map((a, i) => ({ ...a, ordem: i }));
}

/** Plano completo a partir da análise do diagnóstico (problemas críticos primeiro). */
export function montarAtividades(analise: AnaliseDiagnostico, cargaHorariaTotal: number | string | undefined): AtividadeCronograma[] {
  const total = parseCargaHoraria(cargaHorariaTotal);
  const limite = capacidadeAtividades(total);
  const lista = analise.problemas;
  let meio: AtividadeCronograma[];
  if (lista.length > limite) {
    meio = [
      ...lista.slice(0, limite - 1).map(atividadeDoProblema),
      atividadeComplementar(lista.slice(limite - 1).map((p) => ({ nome: p.problema, nivel: p.nivel }))),
    ];
  } else {
    meio = lista.map(atividadeDoProblema);
  }
  return estruturarPlano(
    [atividadeDiagnostico(analise), atividadeDevolutiva(analise), atividadeFerramentas(analise), ...meio, atividadeAnaliseFinal(), atividadeRelatorioFinal()],
    total,
    analise
  );
}

/** @deprecated use estruturarPlano. Mantido para compatibilidade. */
export function garantirPrimeiraEUltima(lista: AtividadeCronograma[], analise?: AnaliseDiagnostico): AtividadeCronograma[] {
  const total = (lista || []).reduce((s, a) => s + horasDe(a), 0);
  return estruturarPlano(lista, total || undefined, analise);
}

// ---------------------------------------------------------------------------------------
// IA: só melhora os TEXTOS (a estrutura, a ordem e as horas são garantidas pelo código)
// ---------------------------------------------------------------------------------------
export interface GeradorIA {
  models: { generateContent: (args: { model: string; contents: any; config?: any }) => Promise<{ text?: string }> };
}

const ID_FERRAMENTAS = '__ferramentas__';

export async function enriquecerComIA(
  ai: GeradorIA | null | undefined,
  atividades: AtividadeCronograma[],
  analise: AnaliseDiagnostico,
  contexto: { tipoEmpresa?: string; nomeEmpresa?: string } = {}
): Promise<{ atividades: AtividadeCronograma[]; usouIA: boolean; erro?: string }> {
  const alvo = analise.problemas.filter((p) => atividades.some((a) => a.idProblema === p.idProblema));
  if (!ai || alvo.length === 0) return { atividades, usouIA: false };

  const dados = alvo.map((p) => ({
    idProblema: p.idProblema,
    problema: p.problema,
    area: p.area,
    criticidade: p.nivel,
    horas: horasDe(atividades.find((a) => a.idProblema === p.idProblema)),
    respostas: p.lacunas.slice(0, 5),
    solucaoCadastrada: p.solucao
      ? { solucao: p.solucao.solucao_recomendada, acoes: p.solucao.acoes_sugeridas, kpis: p.solucao.kpis_sugeridos }
      : undefined,
  }));
  const ferr = atividades.find((a) => papelDaAtividade(a.nome) === 'ferramentas');

  const prompt = `Você é um consultor empresarial sênior (padrão SEBRAE). Com base no diagnóstico${
    contexto.nomeEmpresa ? ` da empresa "${contexto.nomeEmpresa}"` : ''
  }${contexto.tipoEmpresa ? ` (segmento: ${contexto.tipoEmpresa})` : ''}, redija as atividades de consultoria abaixo.

TAREFA 1 — Para CADA problema da lista, escreva uma atividade que SOLUCIONE o problema:
- "nome": título curto e profissional da ação (até 90 caracteres), começando por um verbo no infinitivo.
- "descricao": o que será feito, passo a passo resumido, citando as evidências das respostas do diagnóstico.
- "solucaoProposta": a solução em uma frase.
- "resultadoEsperado": resultado com pelo menos um indicador mensurável (número, % ou prazo) compatível com a carga horária informada.
- "responsavel": "Consultor", "Cliente" ou "Consultor/Cliente".
- Se houver solução cadastrada, use-a como base. Mantenha o mesmo "idProblema".

TAREFA 2 — Inclua também um item com "idProblema": "${ID_FERRAMENTAS}" descrevendo as planilhas e/ou aplicativos de gestão
que serão desenvolvidos para apoiar a solução desses problemas (quais controles, o que cada um mede e quem vai usar).

REGRAS GERAIS: português do Brasil, linguagem objetiva e técnica, sem inventar dados da empresa. NÃO crie atividades de
diagnóstico, devolutiva, análise final ou relatório final (elas já existem no plano).

PROBLEMAS (do mais crítico ao menos crítico):
${JSON.stringify(dados)}
${ferr ? `\nFERRAMENTAS PREVISTAS (${horasDe(ferr)}h): ${ferr.descricao}` : ''}`;

  try {
    const resp = await ai.models.generateContent({
      model: 'gemini-3.6-flash',
      contents: prompt,
      config: {
        responseMimeType: 'application/json',
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.OBJECT,
            properties: {
              idProblema: { type: Type.STRING },
              nome: { type: Type.STRING },
              descricao: { type: Type.STRING },
              solucaoProposta: { type: Type.STRING },
              resultadoEsperado: { type: Type.STRING },
              responsavel: { type: Type.STRING },
            },
            required: ['idProblema', 'nome', 'descricao', 'solucaoProposta', 'resultadoEsperado'],
          },
        },
      },
    });
    const bruto = String(resp?.text || '').replace(/```json|```/g, '').trim();
    const itens: any[] = JSON.parse(bruto.slice(bruto.indexOf('['), bruto.lastIndexOf(']') + 1));
    const porId = new Map(itens.filter((i) => i && i.idProblema).map((i) => [String(i.idProblema), i]));
    const txt = (v: any, atual: any, max = 600) => (typeof v === 'string' && v.trim() ? v.trim().slice(0, max) : atual);

    let aplicados = 0;
    const novas = atividades.map((a) => {
      const papel = papelDaAtividade(a.nome);
      if (papel === 'ferramentas') {
        const i = porId.get(ID_FERRAMENTAS);
        if (!i) return a;
        aplicados++;
        // O nome da atividade fixa não muda (ele identifica a etapa no plano).
        return { ...a, descricao: txt(i.descricao, a.descricao), solucaoProposta: txt(i.solucaoProposta, a.solucaoProposta), resultadoEsperado: txt(i.resultadoEsperado, a.resultadoEsperado) };
      }
      const i = a.idProblema ? porId.get(a.idProblema) : undefined;
      if (!i) return a;
      let nome = txt(i.nome, a.nome, 120);
      // A IA não pode transformar uma atividade de problema em atividade fixa.
      if (papelDaAtividade(nome) !== 'problema') nome = a.nome;
      aplicados++;
      return {
        ...a,
        nome,
        descricao: txt(i.descricao, a.descricao),
        solucaoProposta: txt(i.solucaoProposta, a.solucaoProposta, 300),
        resultadoEsperado: txt(i.resultadoEsperado, a.resultadoEsperado, 400),
        responsavel: txt(i.responsavel, a.responsavel, 40),
      };
    });
    return { atividades: novas, usouIA: aplicados > 0 };
  } catch (e: any) {
    return { atividades, usouIA: false, erro: e?.message || String(e) };
  }
}

// ---------------------------------------------------------------------------------------
// Relatório
// ---------------------------------------------------------------------------------------
export function montarTextosRelatorio(
  analise: AnaliseDiagnostico,
  atividades: AtividadeCronograma[],
  contexto: { nomeEmpresa?: string; tipoEmpresa?: string } = {}
): { objetivo: string; solucoesIndicadas: string; resultadosEsperados: string } {
  const criticos = analise.problemas.filter((p) => p.nivel === 'Crítico');
  const areas = Array.from(new Set(analise.problemas.map((p) => p.area))).slice(0, 5);
  const empresa = contexto.nomeEmpresa ? ` da empresa ${contexto.nomeEmpresa}` : '';
  const totalHoras = atividades.reduce((s, a) => s + horasDe(a), 0);

  const objetivo =
    analise.problemas.length === 0
      ? `Consolidar as boas práticas de gestão identificadas no diagnóstico${empresa} e estruturar controles e indicadores que garantam a continuidade dos resultados.`
      : `Solucionar ${
          criticos.length === 1 ? 'o problema crítico' : criticos.length > 1 ? `os ${criticos.length} problemas críticos` : 'os problemas prioritários'
        } ${criticos.length === 1 ? 'identificado' : 'identificados'} no diagnóstico${empresa}${areas.length ? `, com foco em ${listaTexto(areas)}` : ''}, por meio de ${
          atividades.length
        } atividades (${totalHoras}h), com implantação de ferramentas de gestão e acompanhamento de indicadores. ` +
        `Índice de conformidade no diagnóstico: ${analise.scoreGeral}%.`;

  const problemasTxt = analise.problemas.map((p) => `• [${p.nivel}] ${p.problema} (${p.area})`).join('\n');
  const acoesTxt = atividades
    .map((a, i) => `${i + 1}. ${a.nome} (${a.cargaHoraria})${a.solucaoProposta && a.solucaoProposta !== a.nome ? ` — ${a.solucaoProposta}` : ''}`)
    .join('\n');
  const fortesTxt = analise.pontosFortes.slice(0, 8).map((p) => `• ${p.problema}`).join('\n');

  const solucoesIndicadas = [
    `PROBLEMAS IDENTIFICADOS:\n${problemasTxt || '• Nenhuma lacuna relevante identificada.'}`,
    `SOLUÇÕES / AÇÕES PROPOSTAS:\n${acoesTxt}`,
    fortesTxt ? `PONTOS FORTES:\n${fortesTxt}` : '',
  ]
    .filter(Boolean)
    .join('\n\n');

  const resultadosEsperados = atividades
    .map((a) => a.resultadoEsperado)
    .filter((r): r is string => !!r && !!r.trim())
    .map((r) => `• ${r}`)
    .join('\n');

  return { objetivo, solucoesIndicadas, resultadosEsperados };
}
