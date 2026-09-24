import assert from 'node:assert/strict';
// Simula o Storage do navegador
class FakeStorage { m = new Map<string, string>();
  getItem(k: string) { return this.m.has(k) ? this.m.get(k)! : null; }
  setItem(k: string, v: string) { this.m.set(k, String(v)); }
  removeItem(k: string) { this.m.delete(k); }
  key(i: number) { return Array.from(this.m.keys())[i] ?? null; }
  get length() { return this.m.size; } }
(globalThis as any).Storage = FakeStorage;
const ls = new FakeStorage(); const ss = new FakeStorage();
(globalThis as any).window = { localStorage: ls, sessionStorage: ss };
(globalThis as any).localStorage = ls;

const cs = await import('../src/lib/cryptoStorage');
cs.setStorageUserId('uid_123');
ls.setItem('local_empresas', JSON.stringify([{ id: 'e1', nome: 'Empresa' }]));
ls.setItem('firebase:authUser:xyz', '{"uid":"uid_123"}');
ls.setItem('custom_gemini_api_key', 'AQ.chave');
cs.encryptedLocalStorage.migrateAllToEncrypted();
assert.ok(cs.rawGetItem('local_empresas')!.startsWith('enc:v1:'), 'dados locais criptografados em repouso');
assert.equal(cs.rawGetItem('firebase:authUser:xyz'), '{"uid":"uid_123"}', 'chaves do Firebase não são criptografadas');

// Antes da correção: leitura direta devolvia texto cifrado → JSON.parse quebrava
assert.throws(() => JSON.parse(ls.getItem('local_empresas')!));

cs.instalarLeituraDescriptografada();
assert.deepEqual(JSON.parse(ls.getItem('local_empresas')!), [{ id: 'e1', nome: 'Empresa' }], 'leitura direta descriptografa');
assert.equal(cs.encryptedLocalStorage.getItem('custom_gemini_api_key'), 'AQ.chave', 'encryptedLocalStorage continua funcionando');
ls.setItem('texto_puro', 'abc'); assert.equal(ls.getItem('texto_puro'), 'abc', 'texto puro inalterado');
assert.equal(ls.getItem('nao_existe'), null);

// Outro usuário / antes do login: não consegue abrir → null (e não lixo)
cs.setStorageUserId('outro_usuario');
assert.equal(ls.getItem('local_empresas'), null, 'sem a chave certa devolve null');
cs.setStorageUserId('uid_123');
assert.ok(ls.getItem('local_empresas')!.includes('Empresa'), 'cache limpo ao trocar usuário');

// Migração repetida não criptografa duas vezes
cs.encryptedLocalStorage.migrateAllToEncrypted();
assert.deepEqual(JSON.parse(ls.getItem('local_empresas')!), [{ id: 'e1', nome: 'Empresa' }]);
// sessionStorage não é afetado
ss.setItem('x', 'enc:v1:abc'); assert.equal(ss.getItem('x'), 'enc:v1:abc');
console.log('Todos os testes de cryptoStorage passaram.');
