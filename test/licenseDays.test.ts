import assert from 'node:assert/strict';
import { toJsDate, computeLicenseDaysLeft } from '../src/lib/licenseDays';

const DAY = 86400000;
const ago = (d: number) => Date.now() - d * DAY;

// Formatos de data que apareciam como "NaN dias restantes"
assert.ok(toJsDate({ seconds: Math.floor(ago(5) / 1000), nanoseconds: 0 }));      // cópia do navegador
assert.ok(toJsDate({ _seconds: Math.floor(ago(5) / 1000), _nanoseconds: 0 }));    // Firebase Admin
assert.ok(toJsDate({ toDate: () => new Date(ago(5)) }));                          // Timestamp
assert.ok(toJsDate(new Date(ago(5)).toISOString()));
assert.ok(toJsDate('21/09/2026'));
assert.ok(toJsDate('2026-09-21'));
assert.equal(toJsDate({}), null);
assert.equal(toJsDate('abc'), null);

// Dias restantes nunca NaN
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Teste', diasTeste: 30, dataCadastro: { seconds: Math.floor(ago(10) / 1000) } }), 20);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Teste', diasTeste: '30', dataCadastro: new Date(ago(31)).toISOString() }), -1);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Mensal', dataCadastro: { toDate: () => new Date(ago(0)) } }), 30);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Anual', dataCadastro: {} }), 365);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Definitiva' }), 99999);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Teste', validadeLicenca: new Date(Date.now() + 5 * DAY - 1000).toISOString() }), 5);
assert.equal(computeLicenseDaysLeft({ tipoPlano: 'Teste', diasTeste: 'abc' }), 30);
console.log('OK: dias restantes (14 verificações)');

// --- Planos pagos não podem usar diasTeste (bug: Anual vencia em 30 dias) ---
{
  const { computeLicenseDaysLeft: calc, addDaysIso } = await import('../src/lib/licenseDays');
  const hoje = new Date().toISOString();
  const anual = calc({ tipoPlano: 'Anual', diasTeste: 30, dataCadastro: hoje });
  if (anual < 364) throw new Error(`Anual com diasTeste=30 deveria ter ~365 dias, deu ${anual}`);
  const mensal7 = calc({ tipoPlano: 'Mensal', diasTeste: 7, dataCadastro: hoje });
  if (mensal7 < 29) throw new Error(`Mensal com diasTeste=7 deveria ter ~30 dias, deu ${mensal7}`);
  const teste7 = calc({ tipoPlano: 'Teste', diasTeste: 7, dataCadastro: hoje });
  if (teste7 !== 7) throw new Error(`Teste com 7 dias deveria ter 7, deu ${teste7}`);
  const comValidade = calc({ tipoPlano: 'Anual', diasTeste: 30, validadeLicenca: addDaysIso(365) });
  if (comValidade < 364 || comValidade > 366) throw new Error(`Validade explícita errada: ${comValidade}`);
  console.log('  OK: plano pago ignora diasTeste; validade explícita respeitada');
}
