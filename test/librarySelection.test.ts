import { selectLibraryForDeletion as sel } from '../src/lib/librarySelection';
import assert from 'node:assert/strict';
const problemas = [
  { id: 'p1', descricao_problemas: 'Caixa', area: 'Financeiro', tipoEmpresa: 'Comércio' },
  { id: 'p2', descricao_problemas: 'DRE', area: 'Financeiro', tipoEmpresa: 'Serviços' },
  { id: 'p3', descricao_problemas: 'Insta', area: 'Marketing', tipoEmpresa: 'Comércio' },
  { id: 'p4', descricao_problemas: 'Geral fin', area: ' financeiro ', tipoEmpresa: undefined, tags: ['x'] },
];
const premissas = [
  { id: 'q1', idProblema: 'p1', problema: 'Caixa' },
  { id: 'q2', idProblema: 'p2', problema: 'DRE', tipoEmpresa: 'Serviços' },
  { id: 'q3', idProblema: 'p3', problema: 'Insta' },
  { id: 'q4', problema: 'caixa' },                         // ligada pelo nome
  { id: 'q5', idProblema: 'p4', problema: 'Geral fin' },   // Geral
];
const solucoes = [
  { id: 's1', idProblema: 'p1', problema: 'Caixa', area: 'Financeiro' },
  { id: 's2', idProblema: 'p2', problema: 'DRE', area: 'Financeiro', tipoEmpresa: 'Serviços' },
  { id: 's3', idProblema: 'p3', problema: 'Insta', area: 'Marketing' },
  { id: 's4', idProblema: 'pX', problema: '?', area: 'Financeiro', tipoEmpresa: 'Comércio' }, // solta, mesma seleção
];
const ids = (r: any) => ({ p: r.probs.map((x: any) => x.id), q: r.prems.map((x: any) => x.id), s: r.sols.map((x: any) => x.id) });
const d = { problemas, premissas, solucoes };
assert.deepEqual(ids(sel(d, { area: 'Financeiro', tipo: 'Comércio' })), { p: ['p1'], q: ['q1', 'q4'], s: ['s1', 's4'] });
assert.deepEqual(ids(sel(d, { area: 'financeiro', tipo: 'serviços' })), { p: ['p2'], q: ['q2'], s: ['s2'] });
assert.deepEqual(ids(sel(d, { area: 'Financeiro', tipo: 'Geral' })), { p: ['p4'], q: ['q5'], s: [] });
assert.deepEqual(ids(sel(d, { area: 'Marketing', tipo: 'Comércio' })), { p: ['p3'], q: ['q3'], s: ['s3'] });
assert.deepEqual(ids(sel(d, { area: '', tipo: 'Comércio' })), { p: [], q: [], s: [] });   // sem área: nada
assert.deepEqual(ids(sel(d, { area: 'Financeiro', tipo: 'Geral', tag: 'x' })), { p: ['p4'], q: ['q5'], s: [] });
assert.deepEqual(ids(sel(d, { area: 'Financeiro', tipo: 'Comércio', tag: 'x' })), { p: [], q: [], s: [] });
console.log('OK: 7 cenários de seleção passaram');
