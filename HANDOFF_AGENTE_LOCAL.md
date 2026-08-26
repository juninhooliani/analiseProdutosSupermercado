# Handoff para o Próximo Agente (Claude / Code Agent)

> **Para o agente que assumir este projeto (Claude / Antigravity / Agentic AI):** 
> Este documento registra o resultado das execuções realizadas, o estado atual do repositório e o plano de ação pronto para implementação imediata.
> 
> **Leia antes de iniciar:**
> 1. `agent.md` (Visão geral do estado do projeto)
> 2. `wiki/06_AI_AGENT_RULES.md` (Regras de desenvolvimento e restrições)
> 3. `wiki/02_DATA_MODEL.md` (Modelo de dados e GTIN)

---

## 1. O Que Já Foi Executado & Confirmado

| Tarefa / Etapa | Status | Resultado |
|---|---|---|
| **Higiene do Git** | **CONCLUÍDO** | Repositório `juninhooliani/analiseProdutosSupermercado` tornado **PRIVADO**. Histórico reescrito e limpo de dados sensíveis (sem CPFs/API Keys) publicado via `git push --force-with-lease origin main`. |
| **Exportação e-CPF A1** | **CONCLUÍDO** | Certificado exportado do Windows para `certs/certificado.pfx` com senha local `123456`. Pasta `certs/` devidamente ignorada no Git. |
| **Sonda NFP (`sonda_nfp.py`)** | **CONCLUÍDO** | Executada no Docker. **Resultado:** mTLS puro recebe redirect para a tela de login (`login.aspx` / gov.br). |
| **Definição da Rota do Robô** | **CONCLUÍDO** | A rota aprovada para coleta é **Playwright com o certificado e-CPF injetado no navegador** (`--auto-select-certificate-for-urls`), bypassando a senha/captcha e navegando autenticado. |

---

## 2. Decisão de Arquitetura & Diretrizes de Ingestão

1. **Automação NFP via Playwright**:
   - Usar Playwright com perfil do Chromium contendo o e-CPF A1 instalado.
   - Navegar no portal NFP, realizar o login por Certificado Digital e extrair/baixar os XMLs/HTMLs das notas.
2. **Normalização por GTIN**:
   - O campo **GTIN (Código de Barras EAN)** obtido no XML de NFC-e/NFe é a chave primária de normalização de produtos entre estabelecimentos diferentes (ex: `7891000100103` = Leite Ninho 1L).
3. **Restrição Absoluta**:
   - **NÃO implementar quebra de captcha por OCR (ex: `ddddocr`)**. A rota oficial aprovada é via e-CPF A1 no navegador.

---

## 3. Próximas Tarefas Prontas para Execução

O próximo agente deve focar na implementação das seguintes etapas:

### 🎯 Tarefa 1: Banco de Dados `banco_compras.db` (SQLite + SQLAlchemy/Pydantic)
- Criar a camada de persistência em `core/database.py` e `core/models.py`.
- Implementar as tabelas conforme `wiki/02_DATA_MODEL.md`:
  - `transactions` (Transação Financeira Global)
  - `transaction_items` (Itens de Supermercado/Varejo vinculados à transação, com campo `gtin`)
  - `categories` (Taxonomia de Despesas/Receitas)
  - `accounts` (Contas Bancárias / Cartões)

### 🎯 Tarefa 2: Script de Automação Playwright para NFP (`ingestion/nfp_playwright.py`)
- Construir a automação em Playwright Python usando o certificado e-CPF para autenticar no portal da NFP.
- Extrair/baixar a lista de XMLs das notas fiscais do período e salvar em `XMLs/`.

### 🎯 Tarefa 3: Parser de XMLs NFe/NFCe (`ingestion/xml_parser.py`)
- Ler arquivos XML da pasta `XMLs/`.
- Processar tags de cabeçalho (`emit`, `ide`, `total`) e de produtos (`det -> prod`).
- Extrair GTIN (`cEAN`), descrição do produto (`xProd`), quantidade (`qCom`), valor unitário (`vUnCom`) e valor total.
- Salvar tudo em `banco_compras.db`.

### 🎯 Tarefa 4: Dashboard Web (`web_app/`)
- Construir o Dashboard em Next.js (`web_app/`) conectando com a API/Banco de Dados para exibir gráficos de gastos por categoria, comparação de preços por produto (GTIN) e evolução de inflação pessoal.

---

## 4. Comandos de Apoio

```powershell
# Executar a sonda novamete (se necessário):
docker compose run --rm sonda /certs/certificado.pfx

# Iniciar o ambiente web (Next.js):
cd web_app && npm run dev
```

*Status atualizado em: 26/08/2026.*
