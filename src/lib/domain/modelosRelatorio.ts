/**
 * Modelos de consultoria (plano de trabalho) usados para REPLICAR planos entre clientes
 * do mesmo segmento. Ficam no navegador (e entram no backup) na chave abaixo.
 *
 * Quando um plano é gerado para um segmento que ainda não tem modelo, o app cria
 * automaticamente um "Modelo Padrão" daquele segmento a partir do plano gerado.
 */
import type { AtividadeCronograma } from '../../types/domain';
import { estruturarPlano, horasDe } from './planoAcao';

export interface ModeloRelatorio {
  id: string;
  nome: string;
  categoria: string;
  descricao: string;
  atividades: AtividadeCronograma[];
  padraoAutomatico?: boolean;
  criadoEm?: string;
}

export const MODELOS_STORAGE_KEY = 'custom_modelos_relatorio_v2';
/** Evento disparado quando a lista de modelos muda (a tela de cronograma recarrega). */
export const MODELOS_EVENTO = 'modelos-relatorio-atualizados';

const norm = (s: unknown) =>
  String(s ?? '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();

export function carregarModelos(padroes: ModeloRelatorio[]): ModeloRelatorio[] {
  try {
    const salvo = localStorage.getItem(MODELOS_STORAGE_KEY);
    if (salvo) {
      const lista = JSON.parse(salvo);
      if (Array.isArray(lista) && lista.length) return lista;
    }
  } catch {
    /* usa os padrões */
  }
  return padroes;
}

export function salvarModelos(lista: ModeloRelatorio[]): void {
  try {
    localStorage.setItem(MODELOS_STORAGE_KEY, JSON.stringify(lista));
  } catch (e) {
    console.warn('[Modelos] Não foi possível salvar os modelos:', e);
  }
  try {
    window.dispatchEvent(new CustomEvent(MODELOS_EVENTO));
  } catch {
    /* ambiente sem window */
  }
}

/** Modelo do segmento: categoria igual ao segmento (ignora acentos e maiúsculas). */
export function encontrarModeloDoSegmento(lista: ModeloRelatorio[], segmento?: string): ModeloRelatorio | undefined {
  const alvo = norm(segmento);
  if (!alvo || alvo === 'geral') return undefined;
  return lista.find((m) => norm(m.categoria) === alvo) || lista.find((m) => m.padraoAutomatico && norm(m.nome).includes(alvo));
}

/** Remove dados específicos do cliente (datas, status, evidências, KPIs) para servir de modelo. */
function paraModelo(atividades: AtividadeCronograma[]): AtividadeCronograma[] {
  return atividades.map((a, i) => {
    const { dataInicio, dataFim, evidencias, progressoKPI, metaKPI, ...resto } = a as any;
    return { ...resto, status: 'Pendente', ordem: i } as AtividadeCronograma;
  });
}

export function criarModeloPadrao(segmento: string, atividades: AtividadeCronograma[], cargaHoraria?: string | number): ModeloRelatorio {
  const nomeSegmento = String(segmento || 'Geral').trim() || 'Geral';
  const estrutura = estruturarPlano(atividades, cargaHoraria ?? atividades.reduce((s, a) => s + horasDe(a), 0));
  const horas = estrutura.reduce((s, a) => s + horasDe(a), 0);
  return {
    id: `padrao_${norm(nomeSegmento).replace(/[^a-z0-9]+/g, '_')}_${Date.now()}`,
    nome: `Modelo Padrão — ${nomeSegmento}`,
    categoria: nomeSegmento,
    descricao:
      `Modelo criado automaticamente a partir do diagnóstico (${horas}h, ${estrutura.length} atividades): ` +
      'diagnóstico, devolutiva, ferramentas de gestão, solução dos problemas críticos do segmento, análise final e relatório. ' +
      'Use "Aplicar" ou "Replicar" para reutilizar em outros clientes deste segmento.',
    atividades: paraModelo(estrutura),
    padraoAutomatico: true,
    criadoEm: new Date().toISOString(),
  };
}

/**
 * Garante que o segmento tenha um modelo. Se não houver, cria o "Modelo Padrão" a partir
 * das atividades informadas e salva. Retorna o modelo e se ele foi criado agora.
 */
export function garantirModeloPadrao(
  segmento: string | undefined,
  atividades: AtividadeCronograma[],
  cargaHoraria: string | number | undefined,
  padroes: ModeloRelatorio[]
): { modelo?: ModeloRelatorio; criado: boolean } {
  const seg = String(segmento || '').trim();
  if (!seg || norm(seg) === 'geral' || !atividades?.length) return { criado: false };
  const lista = carregarModelos(padroes);
  const existente = encontrarModeloDoSegmento(lista, seg);
  if (existente) return { modelo: existente, criado: false };
  const modelo = criarModeloPadrao(seg, atividades, cargaHoraria);
  salvarModelos([...lista, modelo]);
  return { modelo, criado: true };
}
