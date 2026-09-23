/**
 * Testes dos formatters brasileiros (Fase 1).
 * Rode com: npx tsx test/formatters.br.test.ts
 */

import {
  cleanDigits,
  formatCNPJ,
  isValidCNPJ,
  formatCPF,
  isValidCPF,
  formatCEP,
} from '../src/lib/formatters/br';

function assert(cond: boolean, msg: string) {
  if (!cond) throw new Error(`FAIL: ${msg}`);
  console.log(`  OK: ${msg}`);
}

console.log('--- formatters/br ---');

assert(cleanDigits('12.345.678/0001-90') === '12345678000190', 'cleanDigits CNPJ');
assert(cleanDigits('111.222.333-44') === '11122233344', 'cleanDigits CPF');
assert(cleanDigits(null) === '', 'cleanDigits null');

assert(formatCNPJ('12345678000190').includes('/'), 'formatCNPJ mascara');
assert(formatCPF('11144477735').includes('-'), 'formatCPF mascara');
assert(formatCEP('01310100') === '01310-100', 'formatCEP');

assert(isValidCNPJ('00000000000000') === false, 'CNPJ zeros inválido');
assert(isValidCNPJ('11111111111111') === false, 'CNPJ repetido inválido');
assert(isValidCPF('00000000000') === false, 'CPF zeros inválido');
assert(isValidCPF('11111111111') === false, 'CPF repetido inválido');

assert(isValidCNPJ('04.252.011/0001-10') === true, 'CNPJ válido 04.252.011/0001-10');
assert(isValidCNPJ('04252011000110') === true, 'CNPJ válido sem máscara');

assert(isValidCPF('529.982.247-25') === true, 'CPF válido 529.982.247-25');
assert(isValidCPF('52998224725') === true, 'CPF válido sem máscara');

console.log('Todos os testes de formatters/br passaram.\n');
