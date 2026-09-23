# Fase 1 — Como integrar no repositório

Os arquivos desta pasta (`src/types`, `src/lib/...`, `test/...`) devem ser **copiados** para a raiz do projeto `ConsultorPro-II`.

## 1. Arquivos já commitados nesta branch

Esta branch já contém os módulos da Fase 1. O `App.tsx` ainda **não** foi alterado (código duplicado permanece até o commit de wiring).

## 2. Alterar o `App.tsx` (mínimo)

No topo do `App.tsx`, **depois** dos imports existentes, adicione os reexports e imports de `./lib/phase1-reexports` (ou imports seletivos).

Depois, **remova** do `App.tsx` os blocos originais correspondentes (interfaces, formatters, constants, etc.).

Veja a lista completa em `docs/APP_TSX_FASE1_PATCH.md` se existir.

## 3. Testes

```bash
npm test
# ou
npx tsx test/formatters.br.test.ts && npx tsx test/domain.respostas.test.ts
```

Atualize o script `test` no package.json para incluir os novos arquivos.

## 4. Checklist

- [ ] App abre sem erro
- [ ] Máscara CNPJ/CPF
- [ ] Diagnóstico carrega respostas
- [ ] npm test / npm run lint
