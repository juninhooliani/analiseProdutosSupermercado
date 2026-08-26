# analiseProdutosSupermercado

Sistema de gestão financeira pessoal com foco em **análise de compras de
supermercado**: histórico de preço por produto, inflação pessoal e comparação de
qual estabelecimento compensa para cada item.

A documentação completa do domínio está em [`wiki/`](wiki/).

## Como funciona

O sistema ingere notas fiscais por quatro caminhos (ver
[`wiki/03_INGESTION_STRATEGY.md`](wiki/03_INGESTION_STRATEGY.md)):

1. **Sessão autenticada por e-CPF** no portal da Nota Fiscal Paulista — rota
   principal, em desenvolvimento. Entrega XML com **GTIN** por item.
2. **Upload manual** de XML / HTML / TXT / PDF.
3. **QR Code** do cupom, para captura no momento da compra.
4. **Extratos bancários** (OFX / CSV), para reconciliação.

Os itens são normalizados e categorizados — por GTIN quando disponível, por
texto (`categorias.py` ou API da OpenAI) como fallback.

## Estado atual

| Módulo | Estado |
|---|---|
| Processamento de cupons TXT -> `compras.xlsx` | **Funcional** |
| Coleta automatizada via e-CPF | Em desenvolvimento (etapa 1: sonda) |
| Banco SQLite unificado | Não iniciado |
| Dashboard web (Next.js) | Estrutura inicializada |

## Uso

### Processamento de cupons em TXT (funcional hoje)

Crie a pasta `cupons_txt/` e jogue os cupons em formato txt nela:

```bash
python Supermercado.py
```

Gera `compras.xlsx` com as colunas: Número do Cupom, Data, Hora, Dia da Semana,
Supermercado, Número, Código de Barras, Descrição, Quantidade, Unidade,
Preço Unitário, Total, Categoria.

A categoria pode vir do arquivo local `categorias.py` ou da API da OpenAI (paga).
Para a segunda opção, crie um `.env` na raiz:

```
OPENAI_API_KEY=sua_chave_aqui
```

> O `.env` está no `.gitignore`. **Nunca** o versione.

### Coleta de notas via e-CPF (em desenvolvimento)

Requer certificado digital **e-CPF A1** e Docker. Ver
[`wiki/08_AMBIENTE_DOCKER.md`](wiki/08_AMBIENTE_DOCKER.md).

```bash
# 1. exportar o certificado do Windows para certs\certificado.pfx
powershell -ExecutionPolicy Bypass -File docker\exportar_certificado.ps1

# 2. construir e rodar a sonda
docker compose build
docker compose run --rm sonda /certs/certificado.pfx
```

### Interface web

```bash
cd web_app
npm run dev     # http://localhost:3000
```

## Sobre captcha

Este projeto **não** implementa contorno automatizado de captcha. A rota
suportada é a autenticada por certificado digital, que além de não violar os
termos do portal entrega um dado melhor (com código de barras). Justificativa
completa em [`wiki/06_AI_AGENT_RULES.md`](wiki/06_AI_AGENT_RULES.md), seção 3.

## Licença

Ver [LICENSE](LICENSE).
