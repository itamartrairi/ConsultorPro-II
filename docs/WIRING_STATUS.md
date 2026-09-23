# Status do wiring (Fase 1 + 2)

## Já no GitHub (`refactor/phase1-utils`)

- [x] Módulos Fase 1 (`src/types`, `src/lib/...`)
- [x] Auth: `useAuth`, `LoginView`
- [x] `main.tsx` com `<AuthProvider>`
- [x] Script `scripts/wire-phase2-app.mjs`
- [x] Testes + `package.json`

## Você precisa rodar localmente (App.tsx ~1MB)

O `App.tsx` é grande demais para alguns pushes automáticos. Aplique o wiring com:

```bash
git fetch origin
git checkout refactor/phase1-utils
node scripts/wire-phase2-app.mjs
npm test
npm run lint
npm run dev
```

O script é **idempotente**: pode rodar mais de uma vez.

### O que o script faz no App.tsx

1. Imports dos módulos Fase 1 (com alias `_` para não colidir com as funções ainda no arquivo)
2. `useAuth()` no início de `App()`
3. Sync de `user` / `loading` com o AuthProvider
4. Substitui a tela de login inline por `<LoginView />`

### Próximo micro-passo (Fase 1 completa)

Remover as funções/constantes duplicadas do `App.tsx` e trocar usos para os imports sem alias `_`.
