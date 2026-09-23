# Refatoração AST (ts-morph)

## Scripts

| Comando | Descrição |
|---------|-----------|
| `npm run ast:inventory` | Lista símbolos de módulo no `App.tsx` (candidatos Fase 1) |
| `npm run ast:remove-phase1` | Remove duplicatas Fase 1 via AST + imports limpos |
| `npm run ast:remove-phase1:dry` | Dry-run (não grava) |

## Fluxo recomendado

```bash
git checkout refactor/phase1-utils

# Auth wiring (string script, estável)
node scripts/wire-phase2-app.mjs

# Remoção segura via AST
npm run ast:remove-phase1

npm test
npm run lint
npm run dev
```

## Por que AST?

O script `remove-phase1-duplicates.mjs` (texto) pode apagar imports demais.
O `ast-remove-phase1.ts` usa **ts-morph** para:

- `getFunction(name).remove()` — escopo real da declaração
- `getInterface` / `getEnum` / variable statements
- Checagem de símbolos críticos (`OperationType`, recharts, Button…)
- Imports sem alias `_`

## Inventário

```bash
npm run ast:inventory
npm run ast:inventory -- --json > /tmp/app-symbols.json
npm run ast:inventory -- --exports-only
```
