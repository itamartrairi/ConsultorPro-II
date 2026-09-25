/**
 * Relatório consolidado: resumo do diagnóstico de TODOS os clientes, agrupado por tipo de
 * negócio (segmento). Usa a mesma análise de criticidade do plano de ação
 * (lib/domain/planoAcao.ts), então os números batem com os relatórios individuais.
 *
 * Para cada cliente considera o diagnóstico MAIS RECENTE que tenha respostas.
 */
import type { Diagnostico, Empresa, Problema, Resposta, Solucao } from '../../types/domain';
import { analisarDiagnostico, grauLacuna, type NivelCriticidade } from './planoAcao';

export interface ResumoCliente {
  empresaId: string;
  cliente: string;
  diagnosticoId: string;
  data: Date | null;
  totalDiagnosticos: number;
  respondidas: number;
  conformidade: number; // 0–100
  criticos: number;
  altos: number;
  moderados: number;
  principaisCriticos: string[];
  /** TODOS os problemas diagnosticados do cliente, do mais ao menos crítico. */
  problemas: ProblemaDiagnosticado[];
  /** Conformidade (0–100) por área deste cliente. */
  areas: Record<string, number>;
}

export interface ProblemaDiagnosticado {
  problema: string;
  area: string;
  nivel: NivelCriticidade;
  /** Perguntas do diagnóstico que evidenciam o problema, com a resposta dada. */
  evidencias: string[];
}

export interface ProblemaRecorrente {
  problema: string;
  area: string;
  clientes: number; // em quantos clientes aparece com lacuna
  comoCritico: number; // em quantos aparece como crítico
  percentual: number; // % dos clientes diagnosticados do segmento
}

export interface ResumoArea {
  area: string;
  conformidadeMedia: number;
  clientesAvaliados: number;
  clientesComLacuna: number;
}

export interface ResumoSegmento {
  segmento: string;
  totalClientes: number;
  clientesDiagnosticados: number;
  clientesSemDiagnostico: string[];
  conformidadeMedia: number;
  conformidadeMin: number;
  conformidadeMax: number;
  criticos: number;
  altos: number;
  moderados: number;
  clientes: ResumoCliente[];
  areas: ResumoArea[];
  recorrentes: ProblemaRecorrente[];
}

export interface RelatorioSegmentos {
  geradoEm: Date;
  /** Tipo de negócio filtrado (vazio = todos). */
  filtroSegmento?: string;
  totalClientes: number;
  totalDiagnosticados: number;
  conformidadeMedia: number;
  segmentos: ResumoSegmento[];
}

const media = (v: number[]) => (v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length) : 0);

function paraData(v: any): Date | null {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (typeof v?.toDate === 'function') return v.toDate();
  if (typeof v === 'object' && typeof v.seconds === 'number') return new Date(v.seconds * 1000);
  const s = String(v);
  const d = /^\d{4}-\d{2}-\d{2}$/.test(s) ? new Date(`${s}T12:00:00`) : new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** Conformidade por área (100 = sem lacunas), ponderada pelo peso das perguntas. */
function conformidadePorArea(respostas: Resposta[]): Record<string, number> {
  const acc: Record<string, { lac: number; max: number }> = {};
  for (const r of respostas) {
    const g = grauLacuna(r);
    if (g === null) continue;
    const area = (r.area || 'Geral').trim() || 'Geral';
    const peso = Number(r.peso) > 0 ? Number(r.peso) : 1;
    acc[area] ||= { lac: 0, max: 0 };
    acc[area].lac += g * peso;
    acc[area].max += 2 * peso;
  }
  const out: Record<string, number> = {};
  for (const [area, v] of Object.entries(acc)) out[area] = v.max ? Math.round((1 - v.lac / v.max) * 100) : 100;
  return out;
}

/** Remove respostas repetidas da mesma pergunta (fica a mais recente respondida). */
function deduplicar(respostas: Resposta[]): Resposta[] {
  const m = new Map<string, Resposta>();
  for (const r of respostas) {
    const chave = `${r.idProblema || r.problema}|${(r.pergunta || '').trim().toLowerCase()}`;
    const atual = m.get(chave);
    if (!atual || (!atual.resposta && r.resposta)) m.set(chave, r);
  }
  return Array.from(m.values());
}

export function montarRelatorioSegmentos(params: {
  empresas: Empresa[];
  diagnosticos: Diagnostico[];
  respostas: Resposta[];
  problemas?: Problema[];
  solucoes?: Solucao[];
  /** Gera só para este tipo de negócio (vazio ou "Todos" = todos). */
  segmento?: string;
}): RelatorioSegmentos {
  const { empresas, diagnosticos, respostas, problemas = [], solucoes = [] } = params;
  const filtro = params.segmento && params.segmento !== 'Todos' ? params.segmento : '';
  const normSeg = (v: string) => v.normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toLowerCase();

  const respostasPorDiag = new Map<string, Resposta[]>();
  for (const r of respostas) {
    if (!r?.diagnosticoId) continue;
    if (!respostasPorDiag.has(r.diagnosticoId)) respostasPorDiag.set(r.diagnosticoId, []);
    respostasPorDiag.get(r.diagnosticoId)!.push(r);
  }

  const segmentoDe = (e?: Empresa, d?: Diagnostico) =>
    (e?.tipoEmpresa || d?.tipoEmpresa || '').trim() || 'Não informado';

  // Empresas conhecidas + empresas citadas só nos diagnósticos
  const empresasMap = new Map<string, Empresa>(empresas.map((e) => [e.id, e]));
  for (const d of diagnosticos) {
    if (d.empresaId && !empresasMap.has(d.empresaId)) {
      empresasMap.set(d.empresaId, { id: d.empresaId, nome: d.nomeEmpresa || 'Cliente sem cadastro', tipoEmpresa: d.tipoEmpresa } as Empresa);
    }
  }

  const porSegmento = new Map<string, { clientes: ResumoCliente[]; semDiagnostico: string[]; total: number }>();
  const garante = (seg: string) => {
    if (!porSegmento.has(seg)) porSegmento.set(seg, { clientes: [], semDiagnostico: [], total: 0 });
    return porSegmento.get(seg)!;
  };
  const analises = new Map<string, ReturnType<typeof analisarDiagnostico>>();

  for (const emp of empresasMap.values()) {
    const diagsDaEmpresa = diagnosticos.filter((d) => d.empresaId === emp.id);
    const comRespostas = diagsDaEmpresa
      .map((d) => ({ d, resps: deduplicar(respostasPorDiag.get(d.id) || []) }))
      .filter((x) => x.resps.some((r) => grauLacuna(r) !== null))
      .sort((a, b) => (paraData(b.d.dataDiagnostico)?.getTime() || 0) - (paraData(a.d.dataDiagnostico)?.getTime() || 0));
    const seg = segmentoDe(emp, comRespostas[0]?.d || diagsDaEmpresa[0]);
    const grupo = garante(seg);
    grupo.total++;
    const nome = (emp as any).nomeFantasia || emp.nome || 'Cliente';
    if (!comRespostas.length) {
      grupo.semDiagnostico.push(nome);
      continue;
    }
    const { d, resps } = comRespostas[0];
    const an = analisarDiagnostico(resps, problemas, solucoes);
    analises.set(d.id, an);
    grupo.clientes.push({
      empresaId: emp.id,
      cliente: nome,
      diagnosticoId: d.id,
      data: paraData(d.dataDiagnostico),
      totalDiagnosticos: diagsDaEmpresa.length,
      respondidas: an.totalRespondidas,
      conformidade: an.scoreGeral,
      criticos: an.problemas.filter((p) => p.nivel === 'Crítico').length,
      altos: an.problemas.filter((p) => p.nivel === 'Alto').length,
      moderados: an.problemas.filter((p) => p.nivel === 'Moderado').length,
      principaisCriticos: an.problemas.filter((p) => p.nivel === 'Crítico').slice(0, 3).map((p) => p.problema),
      problemas: an.problemas.map((p) => ({
        problema: p.problema,
        area: p.area,
        nivel: p.nivel,
        evidencias: p.lacunas.map((l) => `${l.pergunta} (${l.resposta})`),
      })),
      areas: conformidadePorArea(resps),
    });
  }

  const segmentos: ResumoSegmento[] = [];
  for (const [segmento, g] of porSegmento) {
    if (filtro && normSeg(segmento) !== normSeg(filtro)) continue;
    const clientes = g.clientes.sort((a, b) => a.conformidade - b.conformidade || a.cliente.localeCompare(b.cliente));
    const conf = clientes.map((c) => c.conformidade);

    // Áreas: conformidade média entre os clientes que responderam perguntas da área
    const areasAcc = new Map<string, number[]>();
    clientes.forEach((c) => Object.entries(c.areas).forEach(([a, v]) => {
      if (!areasAcc.has(a)) areasAcc.set(a, []);
      areasAcc.get(a)!.push(v);
    }));
    const areas: ResumoArea[] = Array.from(areasAcc, ([area, vals]) => ({
      area,
      conformidadeMedia: media(vals),
      clientesAvaliados: vals.length,
      clientesComLacuna: vals.filter((v) => v < 100).length,
    })).sort((a, b) => a.conformidadeMedia - b.conformidadeMedia || a.area.localeCompare(b.area));

    // Problemas recorrentes no segmento
    const rec = new Map<string, { problema: string; area: string; clientes: Set<string>; criticos: Set<string> }>();
    clientes.forEach((c) => {
      const an = analises.get(c.diagnosticoId)!;
      an.problemas.forEach((p) => {
        const chave = p.problema.trim().toLowerCase();
        if (!rec.has(chave)) rec.set(chave, { problema: p.problema, area: p.area, clientes: new Set(), criticos: new Set() });
        const r = rec.get(chave)!;
        r.clientes.add(c.empresaId);
        if (p.nivel === 'Crítico') r.criticos.add(c.empresaId);
      });
    });
    const recorrentes: ProblemaRecorrente[] = Array.from(rec.values())
      .map((r) => ({
        problema: r.problema,
        area: r.area,
        clientes: r.clientes.size,
        comoCritico: r.criticos.size,
        percentual: clientes.length ? Math.round((r.clientes.size / clientes.length) * 100) : 0,
      }))
      .sort((a, b) => b.comoCritico - a.comoCritico || b.clientes - a.clientes || a.problema.localeCompare(b.problema))
      .slice(0, 10);

    segmentos.push({
      segmento,
      totalClientes: g.total,
      clientesDiagnosticados: clientes.length,
      clientesSemDiagnostico: g.semDiagnostico.sort((a, b) => a.localeCompare(b)),
      conformidadeMedia: media(conf),
      conformidadeMin: conf.length ? Math.min(...conf) : 0,
      conformidadeMax: conf.length ? Math.max(...conf) : 0,
      criticos: clientes.reduce((s, c) => s + c.criticos, 0),
      altos: clientes.reduce((s, c) => s + c.altos, 0),
      moderados: clientes.reduce((s, c) => s + c.moderados, 0),
      clientes,
      areas,
      recorrentes,
    });
  }

  // Segmentos com mais clientes primeiro ("Não informado" por último)
  segmentos.sort(
    (a, b) =>
      Number(a.segmento === 'Não informado') - Number(b.segmento === 'Não informado') ||
      b.clientesDiagnosticados - a.clientesDiagnosticados ||
      a.segmento.localeCompare(b.segmento)
  );

  const todos = segmentos.flatMap((s) => s.clientes);
  return {
    geradoEm: new Date(),
    filtroSegmento: filtro || undefined,
    totalClientes: segmentos.reduce((s, x) => s + x.totalClientes, 0),
    totalDiagnosticados: todos.length,
    conformidadeMedia: media(todos.map((c) => c.conformidade)),
    segmentos,
  };
}

/** Faixa de maturidade pela conformidade (para leitura rápida no relatório). */
export function faixaConformidade(v: number): { rotulo: string; nivel: NivelCriticidade | 'Adequado' } {
  if (v < 40) return { rotulo: 'Crítica', nivel: 'Crítico' };
  if (v < 60) return { rotulo: 'Baixa', nivel: 'Alto' };
  if (v < 80) return { rotulo: 'Intermediária', nivel: 'Moderado' };
  return { rotulo: 'Adequada', nivel: 'Adequado' };
}

/** Tipos de negócio existentes na carteira (para o seletor do relatório). */
export function tiposDeNegocio(empresas: Empresa[], diagnosticos: Diagnostico[] = []): string[] {
  const set = new Map<string, string>();
  const add = (v?: string) => {
    const t = (v || '').trim();
    if (!t) return;
    const k = t.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    if (!set.has(k)) set.set(k, t);
  };
  empresas.forEach((e) => add(e.tipoEmpresa));
  diagnosticos.forEach((d) => add(d.tipoEmpresa));
  const lista = Array.from(set.values()).sort((a, b) => a.localeCompare(b));
  if (empresas.some((e) => !(e.tipoEmpresa || '').trim())) lista.push('Não informado');
  return lista;
}
