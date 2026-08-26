# 01. Visão Geral do Sistema Financeiro Pessoal (Arquitetura do Zero)

## 1. Contexto & Propósito
Este sistema é uma plataforma completa de **Gestão Financeira Pessoal**, desenhada a partir do zero utilizando a filosofia de **OpenWiki** (Base de Conhecimento Estruturada e Otimizada para IA). 

O objetivo do sistema é centralizar todo o ciclo de vida financeiro (Receitas, Despesas Fixas, Despesas Variáveis, Faturas de Cartão, Investimentos e Compras de Supermercado/Varejo) com documentação clara para humanos e agentes inteligentes.

## 2. Princípios de Arquitetura
1. **OpenWiki Specification**: Todo o domínio, modelos de dados, integrações e regras de negócio ficam documentados na pasta `wiki/` de forma legível por LLMs e IAs de código.
2. **Modularidade & Separação de Conceitos**:
   - `core/`: Motor financeiro unificado (Livro Razão, Transações, Contas, Saldo).
   - `ingestion/`: Módulos de entrada de dados (NFs, Extratos de Banco, CSVs, Manual, QR Code).
   - `analytics/`: Motor estatístico, relatórios, métricas de inflação pessoal e gastos.
3. **Persistência Limpa & Escalável**: Banco de dados relacional (SQLite/PostgreSQL) desacoplado da interface gráfica.

## 3. Visão do Fluxo Geral de Dados
```text
[ Fontes de Dados ] (NFs, Extratos, Manual, QR Code)
        │
        ▼
[ Ingestion Layer ] (Parsing & Normalização)
        │
        ▼
[ Core Ledger ] (Transações Unificadas)
        │
        ├──► [ Módulo Varejo / Supermercado (Detalhamento de Itens) ]
        ├──► [ Módulo de Orçamentos & Metas ]
        └──► [ Analytics & Dashboard Web ]
```
