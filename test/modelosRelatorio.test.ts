import assert from 'node:assert/strict';
const mem = new Map<string, string>();
(globalThis as any).localStorage = { getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => mem.set(k, v), removeItem: (k: string) => mem.delete(k) };
let eventos = 0;
(globalThis as any).window = { dispatchEvent: () => { eventos++; } };
(globalThis as any).CustomEvent = class { constructor(public type: string) {} };

const { garantirModeloPadrao, carregarModelos, encontrarModeloDoSegmento, MODELOS_STORAGE_KEY } = await import('../src/lib/domain/modelosRelatorio');
const { montarAtividades, analisarDiagnostico, ATIVIDADE_DIAGNOSTICO, ATIVIDADE_DEVOLUTIVA, ATIVIDADE_FERRAMENTAS, ATIVIDADE_ANALISE_FINAL, ATIVIDADE_RELATORIO_FINAL } = await import('../src/lib/domain/planoAcao');

const padroes = [{ id: 'carc', nome: 'Carcinicultura', categoria: 'Carcinicultura', descricao: '', atividades: [] }];
const an = analisarDiagnostico([{ id: '1', idProblema: 'A', problema: 'Sem fluxo de caixa', pergunta: 'Possui fluxo de caixa?', resposta: 'Não', peso: 2, area: 'Finanças' } as any], [], []);
const plano = montarAtividades(an, 30).map((a) => ({ ...a, dataInicio: '2026-09-01', dataFim: '2026-09-01', status: 'Concluído' as const, evidencias: ['foto-do-cliente.jpg'] }));

// Segmento com modelo → não cria
assert.equal(garantirModeloPadrao('carcinicultura', plano, 30, padroes).criado, false, 'ignora acentos/maiúsculas');
// Segmento sem modelo → cria e salva
const r1 = garantirModeloPadrao('Apicultura', plano, 30, padroes);
assert.equal(r1.criado, true);
assert.equal(r1.modelo!.nome, 'Modelo Padrão — Apicultura');
assert.equal(eventos, 1, 'avisa a tela de modelos');
const salvos = carregarModelos(padroes);
assert.equal(salvos.length, 2, 'mantém os modelos existentes');
const m = encontrarModeloDoSegmento(salvos, 'APICULTURA')!;
assert.deepEqual([m.atividades[0].nome, m.atividades[1].nome, m.atividades[2].nome, m.atividades.at(-2)!.nome, m.atividades.at(-1)!.nome],
  [ATIVIDADE_DIAGNOSTICO, ATIVIDADE_DEVOLUTIVA, ATIVIDADE_FERRAMENTAS, ATIVIDADE_ANALISE_FINAL, ATIVIDADE_RELATORIO_FINAL]);
assert.ok(m.atividades.every((a: any) => a.status === 'Pendente' && !a.dataInicio && !a.evidencias), 'modelo sem dados do cliente');
// Segunda vez → não duplica
assert.equal(garantirModeloPadrao('Apicultura', plano, 30, padroes).criado, false);
assert.equal(JSON.parse(mem.get(MODELOS_STORAGE_KEY)!).length, 2);
// Sem segmento / Geral → não cria
assert.equal(garantirModeloPadrao('Geral', plano, 30, padroes).criado, false);
assert.equal(garantirModeloPadrao(undefined, plano, 30, padroes).criado, false);
console.log('Todos os testes de modelosRelatorio passaram.');
