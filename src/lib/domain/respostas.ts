/**
 * Helpers de respostas do diagnóstico (dedup + storage local).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

import type { Resposta } from '../../types/domain';
import { normalizeAndFormatArea } from './areas';

export function deduplicateRespostas(respostasList: Resposta[]): Resposta[] {
  if (!respostasList || !Array.isArray(respostasList)) return [];
  const map = new Map<string, Resposta>();

  for (const resp of respostasList) {
    if (!resp) continue;
    const diagScope = resp.diagnosticoId ? `${resp.diagnosticoId}::` : '';
    const normArea = resp.area ? `${normalizeAndFormatArea(resp.area).toLowerCase()}::` : '';
    const normQ = (resp.pergunta || '').trim().toLowerCase();
    
    // Identificador único da questão no contexto do diagnóstico e área
    const key = resp.premissaId
      ? `${diagScope}${normArea}prem:${resp.premissaId}`
      : normQ
      ? `${diagScope}${normArea}q:${normQ}`
      : `${diagScope}id:${resp.id}`;
    if (!key) continue;

    if (!map.has(key)) {
      map.set(key, resp);
    } else {
      const existing = map.get(key)!;
      const isAnswered = (r: Resposta) =>
        r.resposta === 'Sim' || r.resposta === 'Parcial' || r.resposta === 'Não';

      const respAnswered = isAnswered(resp);
      const existingAnswered = isAnswered(existing);

      if (respAnswered && !existingAnswered) {
        // Nova resposta preenchida substitui item não respondido
        map.set(key, { ...existing, ...resp });
      } else if (!respAnswered && existingAnswered) {
        // Item existente já respondido tem prioridade sobre template não respondido
        const mergedObs = (resp.observacao && !existing.observacao) ? resp.observacao : existing.observacao;
        map.set(key, { ...resp, ...existing, observacao: mergedObs });
      } else {
        // Ambos respondidos ou ambos pendentes: o item que vem depois na lista
        // representa a alteração mais recente feita pelo consultor no diagnóstico.
        const mergedObs = (resp.observacao !== undefined && resp.observacao !== '')
          ? resp.observacao
          : existing.observacao;
        map.set(key, { ...existing, ...resp, observacao: mergedObs });
      }
    }
  }

  return Array.from(map.values());
}

let cachedAllRespostasMap: Map<string, Resposta> | null = null;

export function loadAllLocalRespostas(): Resposta[] {
  if (cachedAllRespostasMap) {
    return Array.from(cachedAllRespostasMap.values());
  }
  try {
    const saved = localStorage.getItem('local_all_respostas');
    if (saved) {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const deduped = deduplicateRespostas(parsed);
        cachedAllRespostasMap = new Map(deduped.map((r: Resposta) => [r.id, r]));
        return deduped;
      }
    }
    const legacy = localStorage.getItem('local_respostas');
    if (legacy) {
      const parsed = JSON.parse(legacy);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const deduped = deduplicateRespostas(parsed);
        cachedAllRespostasMap = new Map(deduped.map((r: Resposta) => [r.id, r]));
        return deduped;
      }
    }
  } catch (e) {
    console.error('Erro ao carregar respostas locais:', e);
  }
  cachedAllRespostasMap = new Map();
  return [];
}

export function saveAllLocalRespostas(newOrUpdated: Resposta[]): void {
  try {
    if (!newOrUpdated || newOrUpdated.length === 0) return;
    if (!cachedAllRespostasMap) {
      loadAllLocalRespostas();
    }
    const existingList = Array.from(cachedAllRespostasMap!.values());
    const combined = deduplicateRespostas([...existingList, ...newOrUpdated]);
    cachedAllRespostasMap = new Map(combined.map((r: Resposta) => [r.id, r]));

    const all = Array.from(cachedAllRespostasMap.values());
    localStorage.setItem('local_all_respostas', JSON.stringify(all));
    localStorage.setItem('local_respostas', JSON.stringify(all));
  } catch (e) {
    console.error('Erro ao salvar respostas locais:', e);
  }
}

export function getRespostasForDiagnostico(diagId: string): Resposta[] {
  const all = loadAllLocalRespostas();
  return deduplicateRespostas(all.filter((r) => r.diagnosticoId === diagId));
}

export function clearRespostasCache(): void {
  cachedAllRespostasMap = null;
}
