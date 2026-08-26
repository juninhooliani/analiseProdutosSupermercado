# 05. Especificação Técnica da API & Arquitetura de Código

## 1. Organização do Repositório (Do Zero)

```text
/
├── wiki/                     # Base de conhecimento OpenWiki
├── core/                     # Motor financeiro e modelos relacionais (SQLAlchemy / Pydantic)
│   ├── models/               # Schemas de dados
│   ├── ledger.py             # Regras de movimentação bancária
│   └── database.py           # Conexão SQLite / PostgreSQL
├── ingestion/                # Parseadores e importadores
│   ├── xml_parser.py         # Leitor de XML NFe/NFCe
│   ├── qr_parser.py          # Leitor de QR Code de cupons
│   └── ofx_parser.py         # Leitor de extratos bancários
├── analytics/                # Consultas e agregadores de métricas
├── web_app/                  # Interface do Usuário (FastAPI / Next.js / HTML+JS)
└── tests/                    # Suíte de testes unitários e de integração
```

## 2. Padrões de Design
- **Single Responsibility Principle**: Cada parser cuida exclusivamente de um formato de entrada.
- **Imutabilidade de Transações Processadas**: Uma vez gravada, a transação só pode ser editada via estorno ou edição explícita do usuário.
