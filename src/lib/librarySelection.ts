/**
 * Seleção da Biblioteca para o botão "Excluir Problemas": devolve apenas os problemas,
 * perguntas e soluções do Tipo de Empresa + Área escolhidos nos filtros (e da tag, se houver).
 * Perguntas e soluções ligadas a um problema excluído também saem, para não ficarem órfãs.
 */
export interface LibProblema { id: string; descricao_problemas?: string; area?: string; tipoEmpresa?: string; tags?: string[] }
export interface LibPremissa { id: string; idProblema?: string; problema?: string; tipoEmpresa?: string; area?: string }
export interface LibSolucao { id: string; idProblema?: string; problema?: string; tipoEmpresa?: string; area?: string }

const norm = (v: string | undefined | null) => (v || '').trim().toLowerCase();

export function selectLibraryForDeletion<P extends LibProblema, Q extends LibPremissa, S extends LibSolucao>(
  data: { problemas: P[]; premissas: Q[]; solucoes: S[] },
  filters: { area: string; tipo: string; tag?: string }
): { probs: P[]; prems: Q[]; sols: S[] } {
  const { problemas, premissas, solucoes } = data;
  const { area, tipo, tag } = filters;
  if (!area || !tipo) return { probs: [], prems: [], sols: [] };

  const matchTipo = (t?: string | null) => norm(t || 'Geral') === norm(tipo);
  const matchArea = (a?: string | null) => norm(a) === norm(area);
  const findProblem = (idProblema?: string | null, probName?: string | null) => {
    if (idProblema) {
      const found = problemas.find(p => p.id === idProblema);
      if (found) return found;
    }
    if (probName) return problemas.find(p => norm(p.descricao_problemas) === norm(probName));
    return undefined;
  };

  const probs = problemas.filter(p =>
    matchArea(p.area) &&
    matchTipo(p.tipoEmpresa) &&
    (!tag || (p.tags || []).some(t => norm(t) === norm(tag)))
  );
  const probIds = new Set(probs.map(p => p.id));
  const probNames = new Set(probs.map(p => norm(p.descricao_problemas)));
  const linkedToDeleted = (idProblema?: string, probName?: string) =>
    (!!idProblema && probIds.has(idProblema)) ||
    (!idProblema && !!probName && probNames.has(norm(probName)));

  const prems = premissas.filter(q => {
    if (linkedToDeleted(q.idProblema, q.problema)) return true;
    if (tag) return false;
    const prob = findProblem(q.idProblema, q.problema);
    return matchArea(prob?.area ?? q.area) && matchTipo(q.tipoEmpresa || prob?.tipoEmpresa);
  });
  const sols = solucoes.filter(s => {
    if (linkedToDeleted(s.idProblema, s.problema)) return true;
    if (tag) return false;
    const prob = findProblem(s.idProblema, s.problema);
    return matchArea(s.area || prob?.area) && matchTipo(s.tipoEmpresa || prob?.tipoEmpresa);
  });
  return { probs, prems, sols };
}
