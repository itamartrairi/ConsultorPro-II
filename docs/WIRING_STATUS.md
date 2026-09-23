# Status do wiring (Fase 1 + 2)

## Já no GitHub (`refactor/phase1-utils`)

- [x] Módulos Fase 1 (`src/types`, `src/lib/...`)
- [x] Auth: `useAuth`, `LoginView`
- [x] `main.tsx` com `<AuthProvider>`
- [x] `scripts/wire-phase2-app.mjs`
- [x] `scripts/remove-phase1-duplicates.mjs` (**versão segura**)
- [x] Testes + `package.json`

## Aplicar localmente (ordem)

```bash
git fetch origin
git checkout refactor/phase1-utils

# 1) Wiring Auth + imports
node scripts/wire-phase2-app.mjs

# 2) Remover duplicatas (versão segura — preserva imports)
node scripts/remove-phase1-duplicates.mjs

npm test
npm run lint
npm run dev
```

Ambos os scripts são **idempotentes**.

### Proteções do remove script

- Remove funções/constantes **uma a uma** (brace matching)
- **Nunca** apaga imports de recharts, Button, Modal, cryptoStorage, etc.
- Preserva / restaura `enum OperationType`
- Se algo crítico sumir, **aborta sem gravar** o App.tsx

### Depois

```bash
git add src/App.tsx src/main.tsx
git commit -m "refactor: wire App to Phase1 modules and AuthProvider"
git push origin refactor/phase1-utils
```

Abrir PR para `main`.
