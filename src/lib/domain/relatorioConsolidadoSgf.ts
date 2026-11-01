/**
 * Domínio e agregação para o Relatório Consolidado de Diagnósticos:
 * Concatena os diagnósticos por Tipo de Negócio, Código de Contratação SGF e Área Diagnosticada.
 * Classifica cada resposta como Ponto Forte ou Ponto Crítico de acordo com a semântica oficial.
 */

import type { Diagnostico, Empresa, Resposta, Premissa, Problema } from '../../types/domain';
import { grauLacuna } from './planoAcao';
import { normalizeAndFormatArea } from './areas';
import { deduplicateRespostas } from './respostas';

export type ClassificacaoDiagnostico = 'Ponto Forte' | 'Ponto Crítico' | 'Oportunidade de Melhoria' | 'Pendente';

export interface ItemDiagnosticoConsolidado {
  id: string;
  diagnosticoId: string;
  empresaId: string;
  cliente: string;
  tipoNegocio: string;
  codigoSgf: string;
  area: string;
  pergunta: string;
  resposta: 'Sim' | 'Não' | 'Parcial' | string;
  classificacao: ClassificacaoDiagnostico;
  isPontoForte: boolean;
  isPontoCritico: boolean;
  grauLacunaValor: number | null;
  peso: number;
  observacao?: string;
  dataDiagnostico?: Date | null;
}

export interface ComparativoArea {
  area: string;
  totalPerguntas: number;
  pontosFortes: number;
  pontosCriticos: number;
  oportunidades: number;
  percentualPontosFortes: number;
  percentualPontosCriticos: number;
}

export interface ComparativoSegmento {
  segmento: string;
  totalPerguntas: number;
  pontosFortes: number;
  pontosCriticos: number;
  percentualPontosFortes: number;
}

export interface ComparativoSgf {
  codigoSgf: string;
  totalPerguntas: number;
  pontosFortes: number;
  pontosCriticos: number;
  percentualPontosFortes: number;
}

export interface GrupoAreaConsolidada {
  area: string;
  totalItens: number;
  pontosFortes: number;
  pontosCriticos: number;
  itens: ItemDiagnosticoConsolidado[];
}

export interface GrupoSgfConsolidado {
  codigoSgf: string;
  totalItens: number;
  pontosFortes: number;
  pontosCriticos: number;
  areas: GrupoAreaConsolidada[];
}

export interface GrupoTipoNegocioConsolidado {
  tipoNegocio: string;
  totalItens: number;
  pontosFortes: number;
  pontosCriticos: number;
  sgfs: GrupoSgfConsolidado[];
}

export interface FiltrosRelatorioConsolidado {
  tipoNegocio?: string; // 'Todos' ou valor
  codigoSgf?: string;   // 'Todos' ou valor
  area?: string;        // 'Todas' ou valor
  empresaId?: string;   // 'Todas' ou ID
  classificacao?: 'todos' | 'pontos_fortes' | 'pontos_criticos' | 'respondidas';
  buscaTexto?: string;
}

export interface RelatorioConsolidadoSgfData {
  geradoEm: Date;
  filtrosAplicados: FiltrosRelatorioConsolidado;
  totalClientes: number;
  totalDiagnosticos: number;
  totalItens: number;
  totalPontosFortes: number;
  totalPontosCriticos: number;
  totalOportunidades: number;
  totalPendentes: number;
  percentualGeralPontosFortes: number;
  percentualGeralPontosCriticos: number;
  grupos: GrupoTipoNegocioConsolidado[];
  itensFiltrados: ItemDiagnosticoConsolidado[];
  comparativoPorArea: ComparativoArea[];
  comparativoPorSegmento: ComparativoSegmento[];
  comparativoPorSgf: ComparativoSgf[];
  // Listas para alimentação de seletores na UI:
  opcoesTiposNegocio: string[];
  opcoesCodigosSgf: string[];
  opcoesAreas: string[];
  opcoesClientes: { id: string; nome: string }[];
}

function parseDataObj(v: any): Date | null {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v?.toDate === 'function') return v.toDate();
  if (typeof v === 'object' && typeof v.seconds === 'number') return new Date(v.seconds * 1000);
  const s = String(v);
  const d = /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(`${s}T12:00:00`) : new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function normalizarTexto(txt?: string): string {
  return String(txt || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

/**
 * Constrói a lista plana de itens de diagnósticos conectando com as tabelas do sistema
 */
export function extrairItensDiagnosticos(params: {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas: Resposta[];
  premissas?: Premissa[];
  problemas?: Problema[];
}): ItemDiagnosticoConsolidado[] {
  const { empresas, diagnosticos, respostas: rawRespostas, premissas = [], problemas = [] } = params;
  // Desduplica respostas garantindo que as modificações mais recentes prevaleçam
  const respostas = deduplicateRespostas(rawRespostas || []);

  // Mapas por ID e variações de maiúsculas/minúsculas
  const empresasMap = new Map<string, Empresa>();
  empresas.forEach(e => {
    if (e.id) empresasMap.set(e.id, e);
    if ((e as any).nome) empresasMap.set(normalizarTexto((e as any).nome), e);
  });

  const diagnosticosMap = new Map<string, Diagnostico>();
  diagnosticos.forEach(d => {
    if (d.id) diagnosticosMap.set(d.id, d);
  });

  const premissasMap = new Map<string, Premissa>();
  premissas.forEach(p => {
    if (p.id) premissasMap.set(p.id, p);
  });

  const problemasMap = new Map<string, Problema>();
  problemas.forEach(prob => {
    if (prob.id) problemasMap.set(prob.id, prob);
    if (prob.descricao_problemas) problemasMap.set(normalizarTexto(prob.descricao_problemas), prob);
  });

  // Mapear respostas organizadas por diagnosticoId
  const respostasPorDiag = new Map<string, Resposta[]>();
  respostas.forEach(r => {
    if (!r?.diagnosticoId) return;
    if (!respostasPorDiag.has(r.diagnosticoId)) {
      respostasPorDiag.set(r.diagnosticoId, []);
    }
    respostasPorDiag.get(r.diagnosticoId)!.push(r);
  });

  const itens: ItemDiagnosticoConsolidado[] = [];
  const jaAdicionados = new Set<string>();

  // 1. Processar todas as respostas reais existentes
  for (const resp of respostas) {
    if (!resp || !resp.diagnosticoId) continue;

    const diag = diagnosticosMap.get(resp.diagnosticoId);
    const empresaId = diag?.empresaId || (resp as any).empresaId || '';
    const emp = empresasMap.get(empresaId) || empresasMap.get(normalizarTexto(diag?.nomeEmpresa));

    // Identificar perguntas excluídas do diagnóstico
    const excludedIds = diag?.excludedPremissaIds || [];
    const qKey = resp.premissaId || (resp as any).perguntaId || resp.id;
    if (excludedIds.includes(qKey) || (resp.premissaId && excludedIds.includes(resp.premissaId)) || excludedIds.includes(resp.id)) {
      continue;
    }

    const clienteNome =
      (emp as any)?.nomeFantasia ||
      emp?.nome ||
      (emp as any)?.razaoSocial ||
      diag?.nomeEmpresa ||
      diag?.nomeProjeto ||
      (diag as any)?.dadosConsultoria?.razaoSocial ||
      'Cliente';

    // Priorizar tipoEmpresa do diagnóstico específico modificado antes da empresa geral
    const tipoNegocio =
      (diag?.tipoEmpresa || (diag as any)?.dadosConsultoria?.tipoEmpresa || emp?.tipoEmpresa || (emp as any)?.ramoAtividade || '').trim() ||
      'Não informado';

    const codigoSgfRaw =
      (diag?.dadosConsultoria?.codigoSgf || (diag as any)?.codigoSgf || (diag as any)?.codigo_sgf || (resp as any)?.codigoSgf || '').trim();
    const codigoSgf = codigoSgfRaw ? codigoSgfRaw : 'SGF Não Informado';

    // Resolver premissa e problema para enriquecer texto de pergunta e área se estiverem vazios
    const prem = premissasMap.get(resp.premissaId || (resp as any).perguntaId);
    const prob = problemasMap.get(resp.idProblema) || (resp.problema ? problemasMap.get(normalizarTexto(resp.problema)) : undefined);

    const pergunta = (resp.pergunta || prem?.pergunta || '').trim();
    const area = normalizeAndFormatArea(resp.area || prob?.area || prem?.tipoEmpresa || 'Geral');
    const resposta = (resp.resposta || '').trim();
    const observacao = (resp.observacao || '').trim();

    if (!pergunta && !resposta) continue;

    // Chave de desduplicação
    const chaveItem = `${resp.diagnosticoId}_${normalizarTexto(area)}_${resp.premissaId || normalizarTexto(pergunta)}`;
    if (jaAdicionados.has(chaveItem)) continue;
    jaAdicionados.add(chaveItem);

    // Avaliação de lacuna oficial do sistema
    const lac = grauLacuna(resp);

    let classificacao: ClassificacaoDiagnostico;
    let isPontoForte = false;
    let isPontoCritico = false;

    if (lac === 0) {
      classificacao = 'Ponto Forte';
      isPontoForte = true;
    } else if (lac === 1) {
      classificacao = 'Oportunidade de Melhoria';
      isPontoCritico = true;
    } else if (lac === 2) {
      classificacao = 'Ponto Crítico';
      isPontoCritico = true;
    } else {
      classificacao = resposta ? 'Oportunidade de Melhoria' : 'Pendente';
      if (resposta) isPontoCritico = true;
    }

    itens.push({
      id: resp.id || `${resp.diagnosticoId}_${resp.premissaId || Math.random()}`,
      diagnosticoId: resp.diagnosticoId,
      empresaId,
      cliente: clienteNome,
      tipoNegocio,
      codigoSgf,
      area,
      pergunta,
      resposta: resposta || 'Pendente',
      classificacao,
      isPontoForte,
      isPontoCritico,
      grauLacunaValor: lac,
      peso: Number(resp.peso || prem?.peso || 1) > 0 ? Number(resp.peso || prem?.peso || 1) : 1,
      observacao: observacao || undefined,
      dataDiagnostico: parseDataObj(diag?.dataDiagnostico),
    });
  }

  // 2. Se um diagnóstico tem premissas cadastradas para suas áreas que ainda não tinham resposta
  // ou se o diagnóstico não tinha respostas criadas ainda, gerar os itens pendentes a partir das premissas
  for (const diag of diagnosticos) {
    const respsDoDiag = respostasPorDiag.get(diag.id) || [];
    const emp = empresasMap.get(diag.empresaId || '') || empresasMap.get(normalizarTexto(diag.nomeEmpresa));
    const clienteNome =
      (emp as any)?.nomeFantasia ||
      emp?.nome ||
      (emp as any)?.razaoSocial ||
      diag.nomeEmpresa ||
      diag.nomeProjeto ||
      (diag as any)?.dadosConsultoria?.razaoSocial ||
      'Cliente';

    const tipoNegocio =
      (diag.tipoEmpresa || (diag as any)?.dadosConsultoria?.tipoEmpresa || emp?.tipoEmpresa || (emp as any)?.ramoAtividade || '').trim() ||
      'Não informado';

    const codigoSgfRaw =
      (diag.dadosConsultoria?.codigoSgf || (diag as any)?.codigoSgf || (diag as any)?.codigo_sgf || '').trim();
    const codigoSgf = codigoSgfRaw ? codigoSgfRaw : 'SGF Não Informado';

    const areasDiag = diag.areasDiagnostico || [];
    const excludedIds = diag.excludedPremissaIds || [];

    // Se o diagnóstico tem poucas respostas ou nenhuma, buscar nas premissas das áreas do diagnóstico
    if (respsDoDiag.length === 0 && premissas.length > 0) {
      for (const p of premissas) {
        if (excludedIds.includes(p.id)) continue;
        const prob = problemasMap.get(p.idProblema) || problemasMap.get(normalizarTexto(p.problema));
        const areaPremissa = prob?.area || 'Geral';
        const normArea = normalizeAndFormatArea(areaPremissa);

        // Se o diagnóstico especificou áreas, respeitar
        if (areasDiag.length > 0 && !areasDiag.some(a => normalizeAndFormatArea(a).toLowerCase() === normArea.toLowerCase())) {
          continue;
        }

        const pergunta = (p.pergunta || '').trim();
        if (!pergunta) continue;

        const chaveItem = `${diag.id}_${normalizarTexto(normArea)}_${p.id || normalizarTexto(pergunta)}`;
        if (jaAdicionados.has(chaveItem)) continue;
        jaAdicionados.add(chaveItem);

        itens.push({
          id: `item_prem_${diag.id}_${p.id}`,
          diagnosticoId: diag.id,
          empresaId: diag.empresaId,
          cliente: clienteNome,
          tipoNegocio,
          codigoSgf,
          area: normArea,
          pergunta,
          resposta: 'Pendente',
          classificacao: 'Pendente',
          isPontoForte: false,
          isPontoCritico: false,
          grauLacunaValor: null,
          peso: Number(p.peso) > 0 ? Number(p.peso) : 1,
          observacao: undefined,
          dataDiagnostico: parseDataObj(diag.dataDiagnostico),
        });
      }
    }
  }

  return itens;
}

/**
 * Monta o relatório consolidado completo por Tipo de Negócio, Código SGF e Área Diagnosticada
 */
export function montarRelatorioConsolidadoSgf(params: {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas: Resposta[];
  premissas?: Premissa[];
  problemas?: Problema[];
  filtros?: FiltrosRelatorioConsolidado;
}): RelatorioConsolidadoSgfData {
  const todosItens = extrairItensDiagnosticos(params);
  const filtros = params.filtros || {};

  // Listas únicas para os seletores
  const tiposSet = new Set<string>();
  const sgfSet = new Set<string>();
  const areasSet = new Set<string>();
  const clientesMap = new Map<string, string>();

  // Também popular seletores com empresas e diagnósticos mesmo se não tiverem perguntas
  params.empresas.forEach(e => {
    const nome = (e as any).nomeFantasia || e.nome || (e as any).razaoSocial || 'Cliente';
    if (e.id) clientesMap.set(e.id, nome);
    const tipo = (e.tipoEmpresa || (e as any).ramoAtividade || '').trim();
    if (tipo) tiposSet.add(tipo);
  });

  params.diagnosticos.forEach(d => {
    const sgf = (d.dadosConsultoria?.codigoSgf || (d as any).codigoSgf || '').trim();
    if (sgf) sgfSet.add(sgf);
    const tipo = (d.tipoEmpresa || '').trim();
    if (tipo) tiposSet.add(tipo);
    if (d.empresaId && !clientesMap.has(d.empresaId)) {
      clientesMap.set(d.empresaId, d.nomeEmpresa || 'Cliente');
    }
  });

  todosItens.forEach(it => {
    if (it.tipoNegocio) tiposSet.add(it.tipoNegocio);
    if (it.codigoSgf) sgfSet.add(it.codigoSgf);
    if (it.area) areasSet.add(it.area);
    if (it.empresaId && !clientesMap.has(it.empresaId)) {
      clientesMap.set(it.empresaId, it.cliente);
    }
  });

  const opcoesTiposNegocio = Array.from(tiposSet).sort((a, b) => a.localeCompare(b));
  const opcoesCodigosSgf = Array.from(sgfSet).sort((a, b) => {
    if (a === 'SGF Não Informado') return 1;
    if (b === 'SGF Não Informado') return -1;
    return a.localeCompare(b);
  });
  const opcoesAreas = Array.from(areasSet).sort((a, b) => a.localeCompare(b));
  const opcoesClientes = Array.from(clientesMap.entries())
    .map(([id, nome]) => ({ id, nome }))
    .sort((a, b) => a.nome.localeCompare(b.nome));

  // Aplicação dos filtros
  const buscaNorm = normalizarTexto(filtros.buscaTexto);
  const tipoFiltro = filtros.tipoNegocio && filtros.tipoNegocio !== 'Todos' ? normalizarTexto(filtros.tipoNegocio) : '';
  const sgfFiltro = filtros.codigoSgf && filtros.codigoSgf !== 'Todos' ? normalizarTexto(filtros.codigoSgf) : '';
  const areaFiltro = filtros.area && filtros.area !== 'Todas' ? normalizarTexto(filtros.area) : '';
  const empresaFiltro = filtros.empresaId && filtros.empresaId !== 'Todas' ? filtros.empresaId : '';

  const itensFiltrados = todosItens.filter(it => {
    if (tipoFiltro && normalizarTexto(it.tipoNegocio) !== tipoFiltro) return false;
    if (sgfFiltro && normalizarTexto(it.codigoSgf) !== sgfFiltro) return false;
    if (areaFiltro && normalizarTexto(it.area) !== areaFiltro) return false;
    if (empresaFiltro && it.empresaId !== empresaFiltro) return false;

    if (filtros.classificacao === 'pontos_fortes' && !it.isPontoForte) return false;
    if (filtros.classificacao === 'pontos_criticos' && !it.isPontoCritico) return false;
    if (filtros.classificacao === 'respondidas' && it.classificacao === 'Pendente') return false;

    if (buscaNorm) {
      const match =
        normalizarTexto(it.cliente).includes(buscaNorm) ||
        normalizarTexto(it.pergunta).includes(buscaNorm) ||
        normalizarTexto(it.resposta).includes(buscaNorm) ||
        normalizarTexto(it.area).includes(buscaNorm) ||
        normalizarTexto(it.codigoSgf).includes(buscaNorm) ||
        normalizarTexto(it.observacao).includes(buscaNorm);
      if (!match) return false;
    }

    return true;
  });

  // Estatísticas gerais
  const totalItens = itensFiltrados.length;
  const totalPontosFortes = itensFiltrados.filter(i => i.isPontoForte).length;
  const totalPontosCriticos = itensFiltrados.filter(i => i.classificacao === 'Ponto Crítico').length;
  const totalOportunidades = itensFiltrados.filter(i => i.classificacao === 'Oportunidade de Melhoria').length;
  const totalPendentes = itensFiltrados.filter(i => i.classificacao === 'Pendente').length;

  const totalAvaliados = totalPontosFortes + totalPontosCriticos + totalOportunidades;
  const percentualGeralPontosFortes = totalAvaliados ? Math.round((totalPontosFortes / totalAvaliados) * 100) : 0;
  const percentualGeralPontosCriticos = totalAvaliados ? Math.round(((totalPontosCriticos + totalOportunidades) / totalAvaliados) * 100) : 0;

  const clientesUnicos = new Set(itensFiltrados.map(i => i.empresaId).filter(Boolean));
  const diagsUnicos = new Set(itensFiltrados.map(i => i.diagnosticoId).filter(Boolean));

  // Comparativo por Área
  const areasMap = new Map<string, { fortes: number; criticos: number; oport: number; total: number }>();
  itensFiltrados.forEach(it => {
    if (!areasMap.has(it.area)) {
      areasMap.set(it.area, { fortes: 0, criticos: 0, oport: 0, total: 0 });
    }
    const acc = areasMap.get(it.area)!;
    acc.total++;
    if (it.isPontoForte) acc.fortes++;
    else if (it.classificacao === 'Ponto Crítico') acc.criticos++;
    else if (it.classificacao === 'Oportunidade de Melhoria') acc.oport++;
  });

  const comparativoPorArea: ComparativoArea[] = Array.from(areasMap.entries())
    .map(([area, v]) => {
      const avaliados = v.fortes + v.criticos + v.oport;
      return {
        area,
        totalPerguntas: v.total,
        pontosFortes: v.fortes,
        pontosCriticos: v.criticos + v.oport,
        oportunidades: v.oport,
        percentualPontosFortes: avaliados ? Math.round((v.fortes / avaliados) * 100) : 0,
        percentualPontosCriticos: avaliados ? Math.round(((v.criticos + v.oport) / avaliados) * 100) : 0,
      };
    })
    .sort((a, b) => b.pontosCriticos - a.pontosCriticos || a.area.localeCompare(b.area));

  // Comparativo por Tipo de Negócio (Segmento)
  const segmentoMap = new Map<string, { fortes: number; criticos: number; total: number }>();
  itensFiltrados.forEach(it => {
    if (!segmentoMap.has(it.tipoNegocio)) {
      segmentoMap.set(it.tipoNegocio, { fortes: 0, criticos: 0, total: 0 });
    }
    const acc = segmentoMap.get(it.tipoNegocio)!;
    acc.total++;
    if (it.isPontoForte) acc.fortes++;
    else if (it.isPontoCritico) acc.criticos++;
  });

  const comparativoPorSegmento: ComparativoSegmento[] = Array.from(segmentoMap.entries())
    .map(([segmento, v]) => {
      const avaliados = v.fortes + v.criticos;
      return {
        segmento,
        totalPerguntas: v.total,
        pontosFortes: v.fortes,
        pontosCriticos: v.criticos,
        percentualPontosFortes: avaliados ? Math.round((v.fortes / avaliados) * 100) : 0,
      };
    })
    .sort((a, b) => b.totalPerguntas - a.totalPerguntas);

  // Comparativo por Código SGF
  const sgfMap = new Map<string, { fortes: number; criticos: number; total: number }>();
  itensFiltrados.forEach(it => {
    if (!sgfMap.has(it.codigoSgf)) {
      sgfMap.set(it.codigoSgf, { fortes: 0, criticos: 0, total: 0 });
    }
    const acc = sgfMap.get(it.codigoSgf)!;
    acc.total++;
    if (it.isPontoForte) acc.fortes++;
    else if (it.isPontoCritico) acc.criticos++;
  });

  const comparativoPorSgf: ComparativoSgf[] = Array.from(sgfMap.entries())
    .map(([codigoSgf, v]) => {
      const avaliados = v.fortes + v.criticos;
      return {
        codigoSgf,
        totalPerguntas: v.total,
        pontosFortes: v.fortes,
        pontosCriticos: v.criticos,
        percentualPontosFortes: avaliados ? Math.round((v.fortes / avaliados) * 100) : 0,
      };
    })
    .sort((a, b) => b.totalPerguntas - a.totalPerguntas);

  // Agrupamento em 3 níveis: Tipo de Negócio -> Código SGF -> Área
  const estruturaPorTipo = new Map<string, Map<string, Map<string, ItemDiagnosticoConsolidado[]>>>();

  itensFiltrados.forEach(it => {
    if (!estruturaPorTipo.has(it.tipoNegocio)) {
      estruturaPorTipo.set(it.tipoNegocio, new Map());
    }
    const porSgf = estruturaPorTipo.get(it.tipoNegocio)!;

    if (!porSgf.has(it.codigoSgf)) {
      porSgf.set(it.codigoSgf, new Map());
    }
    const porArea = porSgf.get(it.codigoSgf)!;

    if (!porArea.has(it.area)) {
      porArea.set(it.area, []);
    }
    porArea.get(it.area)!.push(it);
  });

  const grupos: GrupoTipoNegocioConsolidado[] = [];

  for (const [tipoNegocio, porSgf] of estruturaPorTipo) {
    const sgfs: GrupoSgfConsolidado[] = [];
    let totalItensTipo = 0;
    let fortesTipo = 0;
    let criticosTipo = 0;

    for (const [codigoSgf, porArea] of porSgf) {
      const areas: GrupoAreaConsolidada[] = [];
      let totalItensSgf = 0;
      let fortesSgf = 0;
      let criticosSgf = 0;

      for (const [area, itensArea] of porArea) {
        // Ordenar itens da área: pontos críticos primeiro (para destaque), depois por cliente
        itensArea.sort((a, b) => {
          if (a.isPontoCritico && !b.isPontoCritico) return -1;
          if (!a.isPontoCritico && b.isPontoCritico) return 1;
          return a.cliente.localeCompare(b.cliente) || a.pergunta.localeCompare(b.pergunta);
        });

        const fArea = itensArea.filter(i => i.isPontoForte).length;
        const cArea = itensArea.filter(i => i.isPontoCritico).length;

        areas.push({
          area,
          totalItens: itensArea.length,
          pontosFortes: fArea,
          pontosCriticos: cArea,
          itens: itensArea,
        });

        totalItensSgf += itensArea.length;
        fortesSgf += fArea;
        criticosSgf += cArea;
      }

      // Ordenar áreas por quantidade de pontos críticos
      areas.sort((a, b) => b.pontosCriticos - a.pontosCriticos || a.area.localeCompare(b.area));

      sgfs.push({
        codigoSgf,
        totalItens: totalItensSgf,
        pontosFortes: fortesSgf,
        pontosCriticos: criticosSgf,
        areas,
      });

      totalItensTipo += totalItensSgf;
      fortesTipo += fortesSgf;
      criticosTipo += criticosSgf;
    }

    sgfs.sort((a, b) => b.totalItens - a.totalItens);

    grupos.push({
      tipoNegocio,
      totalItens: totalItensTipo,
      pontosFortes: fortesTipo,
      pontosCriticos: criticosTipo,
      sgfs,
    });
  }

  grupos.sort((a, b) => b.totalItens - a.totalItens || a.tipoNegocio.localeCompare(b.tipoNegocio));

  return {
    geradoEm: new Date(),
    filtrosAplicados: filtros,
    totalClientes: clientesUnicos.size,
    totalDiagnosticos: diagsUnicos.size,
    totalItens,
    totalPontosFortes,
    totalPontosCriticos,
    totalOportunidades,
    totalPendentes,
    percentualGeralPontosFortes,
    percentualGeralPontosCriticos,
    grupos,
    itensFiltrados,
    comparativoPorArea,
    comparativoPorSegmento,
    comparativoPorSgf,
    opcoesTiposNegocio,
    opcoesCodigosSgf,
    opcoesAreas,
    opcoesClientes,
  };
}
