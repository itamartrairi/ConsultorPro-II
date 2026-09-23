/**
 * Helpers de respostas do diagnóstico (dedup + storage local).
 * Extraído do App.tsx (Fase 1 da refatoração).
 */

import type { Resposta } from '../../types/domain';

export function deduplicateRespostas(respostasList: Resposta[]): Resposta[] {
  if (!respostasList || !Array.isArray(respostasList)) return [];
  const map = new Map<string, Resposta>();

  for (const resp of respostasList) {
    if (!resp) continue;
    const normQ = (resp.pergunta || '').trim().toLowerCase();
    const key = normQ ? `q:${normQ}` : resp.premissaId ? `id:${resp.premissaId}` : resp.id;
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
        map.set(key, resp);
      } else if (respAnswered && existingAnswered) {
        if (resp.observacao && !existing.observacao) {
          map.set(key, resp);
        } else if ((resp.score || 0) > (existing.score || 0)) {
          map.set(key, resp);
        }
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
        cachedAllRespostasMap = new Map(parsed.map((r: Resposta) => [r.id, r]));
        return parsed;
      }
    }
    const legacy = localStorage.getItem('local_respostas');
    if (legacy) {
      const parsed = JSON.parse(legacy);
      if (Array.isArray(parsed) && parsed.length > 0) {
        cachedAllRespostasMap = new Map(parsed.map((r: Resposta) => [r.id, r]));
        return parsed;
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
    newOrUpdated.forEach((r) => {
      if (r?.id) cachedAllRespostasMap!.set(r.id, r);
    });
    const all = Array.from(cachedAllRespostasMap!.values());
    localStorage.setItem('local_all_respostas', JSON.stringify(all));
    localStorage.setItem('local_respostas', JSON.stringify(all));
  } catch (e) {
    console.error('Erro ao salvar respostas locais:', e);
  }
}

export function getRespostasForDiagnostico(diagId: string): Resposta[] {
  const all = loadAllLocalRespostas();
  return all.filter((r) => r.diagnosticoId === diagId);
}

export function clearRespostasCache(): void {
  cachedAllRespostasMap = null;
}
