/**
 * Testes de deduplicateRespostas (Fase 1).
 * Rode com: npx tsx test/domain.respostas.test.ts
 */

import { deduplicateRespostas } from '../src/lib/domain/respostas';
import type { Resposta } from '../src/types/domain';

function assert(cond: boolean, msg: string) {
  if (!cond) throw new Error(`FAIL: ${msg}`);
  console.log(`  OK: ${msg}`);
}

function makeResp(partial: Partial<Resposta> & { id: string; pergunta: string }): Resposta {
  return {
    diagnosticoId: 'd1',
    premissaId: '',
    idProblema: 'p1',
    problema: 'Problema',
    peso: 1,
    area: 'Financeiro',
    observacao: '',
    score: 0,
    ...partial,
  };
}

console.log('--- domain/respostas ---');

assert(deduplicateRespostas([]).length === 0, 'lista vazia');
assert(deduplicateRespostas(null as any).length === 0, 'null seguro');

const a = makeResp({ id: '1', pergunta: 'Possui controle?', resposta: '', score: 0 });
const b = makeResp({
  id: '2',
  pergunta: 'Possui controle?',
  resposta: 'Sim',
  score: 5,
  observacao: 'ok',
});
const dedup = deduplicateRespostas([a, b]);
assert(dedup.length === 1, 'mesma pergunta → 1 item');
assert(dedup[0].resposta === 'Sim', 'prioriza respondida');
assert(dedup[0].id === '2', 'fica a respondida');

const c = makeResp({ id: '3', pergunta: 'Outra pergunta?', resposta: 'Não', score: 1 });
const dedup2 = deduplicateRespostas([a, b, c]);
assert(dedup2.length === 2, 'perguntas diferentes → 2 itens');

// Cenário crítico: Consultor altera a resposta de 'Sim' para 'Não' (grau de lacuna maior / ponto crítico)
const antesSim = makeResp({ id: '10', pergunta: 'Tem controle?', resposta: 'Sim', score: 2, observacao: 'antes' });
const depoisNao = makeResp({ id: '10', pergunta: 'Tem controle?', resposta: 'Não', score: 0, observacao: 'precisa implantar' });
const alteradoParaNao = deduplicateRespostas([antesSim, depoisNao]);
assert(alteradoParaNao.length === 1, 'alteração Sim para Não mantém 1 item');
assert(alteradoParaNao[0].resposta === 'Não', 'alteração de Sim para Não deve prevalecer');
assert(alteradoParaNao[0].observacao === 'precisa implantar', 'observação atualizada prevalece');

// Cenário inverso: Consultor altera de 'Não' para 'Sim' (resolvido / ponto forte)
const depoisSim = makeResp({ id: '10', pergunta: 'Tem controle?', resposta: 'Sim', score: 2, observacao: 'implantado com sucesso' });
const alteradoParaSim = deduplicateRespostas([depoisNao, depoisSim]);
assert(alteradoParaSim.length === 1, 'alteração Não para Sim mantém 1 item');
assert(alteradoParaSim[0].resposta === 'Sim', 'alteração de Não para Sim deve prevalecer');

console.log('Todos os testes de domain/respostas passaram.\n');
