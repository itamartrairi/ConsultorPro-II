# O que fazer agora (passo a passo simples)

A branch `refactor/phase1-utils` já tem os módulos e scripts.
Você só precisa **aplicar** no `App.tsx` e testar.

## Opção A — No computador (recomendado)

### 1. Abrir o terminal na pasta do projeto

Se o projeto já está baixado:

```bash
cd caminho/para/ConsultorPro-II
```

Se ainda não baixou:

```bash
git clone https://github.com/itamartrairi/ConsultorPro-II.git
cd ConsultorPro-II
```

### 2. Entrar na branch certa

```bash
git fetch origin
git checkout refactor/phase1-utils
npm install
```

### 3. Aplicar a refatoração (copie e cole)

```bash
node scripts/wire-phase2-app.mjs
npm run ast:remove-phase1
```

### 4. Testar

```bash
npm test
npm run lint
npm run dev
```

Abra o endereço que aparecer no terminal (ex.: http://localhost:3000).
Teste login e se o app abre.

### 5. Enviar para o GitHub

```bash
git add src/App.tsx src/main.tsx
git commit -m "refactor: wire App to Phase1 modules and AuthProvider"
git push origin refactor/phase1-utils
```

### 6. Abrir o Pull Request

No navegador:

https://github.com/itamartrairi/ConsultorPro-II/compare/main...refactor/phase1-utils

Clique em **Create pull request** → **Create pull request**.

---

## Opção B — Só com o site do GitHub

1. Abra: https://github.com/itamartrairi/ConsultorPro-II/tree/refactor/phase1-utils
2. Se não souber usar terminal, peça a alguém com Git instalado para rodar a **Opção A**.
3. Ou use o **GitHub Desktop** (programa com botões):
   - Baixe: https://desktop.github.com/
   - File → Clone repository → escolha ConsultorPro-II
   - Current branch → `refactor/phase1-utils`
   - Depois ainda precisa do terminal para os comandos do passo 3.

---

## Se der erro

- `npm: command not found` → instale Node.js: https://nodejs.org/
- `git: command not found` → instale Git: https://git-scm.com/
- Erro no `ast:remove-phase1` → rode só:
  ```bash
  node scripts/wire-phase2-app.mjs
  npm run dev
  ```
  (pelo menos o login novo já funciona)

---

## Resumo em 1 linha

```bash
git checkout refactor/phase1-utils && npm install && node scripts/wire-phase2-app.mjs && npm run ast:remove-phase1 && npm run dev
```
