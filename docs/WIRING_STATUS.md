# Status do wiring (Fase 1 + 2)

## Já no GitHub (`refactor/phase1-utils`)

- [x] Módulos Fase 1 (`src/types`, `src/lib/...`)
- [x] Auth: `useAuth`, `LoginView`
- [x] `main.tsx` com `<AuthProvider>`
- [x] `scripts/wire-phase2-app.mjs`
- [x] `scripts/remove-phase1-duplicates.mjs`
- [x] Testes + `package.json`

## Aplicar localmente (ordem)

```bash
git fetch origin
git checkout refactor/phase1-utils

# 1) Wiring Auth + imports
node scripts/wire-phase2-app.mjs

# 2) Remover duplicatas e usar módulos Fase 1
node scripts/remove-phase1-duplicates.mjs

npm test
npm run lint
npm run dev
```

Ambos os scripts são **idempotentes**.

### O que o remove-phase1-duplicates faz

1. Troca imports com alias `_` por imports reais + reexports
2. Remove do `App.tsx`:
   - enum `Type`
   - formatters BR / dates
   - sanitize, oklch, logo
   - interfaces de domínio
   - AREAS, CONSULTORIA_AREAS, TIPOS_EMPRESA, AREAS_ORDER
   - deduplicateRespostas, extractAndParseJSON, load/save respostas
3. Reduz ~30–40 KB do App (o restante continua sendo UI/views)

### Depois

```bash
git add src/App.tsx src/main.tsx
git commit -m "refactor: wire App to Phase1 modules and AuthProvider"
git push origin refactor/phase1-utils
```

Abrir PR para `main`.
