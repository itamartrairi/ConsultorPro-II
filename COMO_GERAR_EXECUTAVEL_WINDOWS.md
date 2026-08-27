# Como Gerar o Instalador Executável (.exe) para Windows

O projeto já está 100% configurado com o **Electron** e **Electron Builder** para gerar instaladores nativos do Windows.

---

## 🚀 Passo a Passo no seu Computador (Windows)

### 1. Baixar o projeto / Clonar do GitHub
Se você baixou o ZIP do projeto ou clonou o repositório no seu computador:

1. Abra a pasta do projeto no terminal (Prompt de Comando, PowerShell ou VS Code).
2. Instale as dependências executando:
   ```bash
   npm install
   ```

### 2. Gerar o Instalador `.exe`
Execute o comando de build para Windows:
```bash
npm run electron:build:win
```

### 3. Onde encontrar o arquivo gerado
Após a finalização do comando (leva cerca de 1 a 2 minutos na primeira vez), os arquivos executáveis estarão dentro da pasta criada:
📁 **`release/`**

Lá você encontrará:
- **`Gestor Consultor Pro Setup 1.0.0.exe`**: Instalador completo com assistente de instalação, criação de atalho na Área de Trabalho e Menu Iniciar.
- **`Gestor Consultor Pro 1.0.0.exe`**: Versão portátil (Portable) que roda direto de um pendrive sem precisar instalar.

---

## 🧪 Como Testar no Modo Desktop antes de Gerar o Executável
Para abrir a janela do aplicativo no modo desktop para testes:
```bash
npm run electron:dev
```
