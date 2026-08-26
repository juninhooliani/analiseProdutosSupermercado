# 06. Regras de Negócio e Contexto para Agentes de IA

## 1. Diretrizes para a IA Coding Assistant
- Sempre consulte este diretório `wiki/` antes de criar novos schemas ou refatorar o código existente.
- Mantenha a separação entre a lógica do *Core Financeiro* e as rotinas específicas de *Ingestão de Notas Fiscais*.
- Caso detecte um produto novo sem categoria em `categorias.py`, solicite sugestão de categorização ou adicione ao dicionário padrão.

## 2. Padrões de Código e Manipulação de Dados
- Todo valor financeiro deve usar números decimais com 2 casas de precisão (`Decimal` preferencialmente a `float`).
- Datas devem seguir rigorosamente o formato ISO-8601 (`YYYY-MM-DD`).
- Scripts novos devem fornecer suporte tanto para execuções em linha de comando (CLI) quanto importações em API/Web App.
- **GTIN tem precedência sobre nome normalizado** na identificação de produtos. Ver `03_INGESTION_STRATEGY.md`, seção 4.

## 3. Restrição: quebra de captcha está fora de escopo

O projeto **não** implementa contorno automatizado de captcha. Uma versão
anterior tinha um robô com `ddddocr` (`extrator_sefaz.py`) apontando para a
consulta pública da NFC-e; ele foi removido em favor da rota do e-CPF.

Motivos, em ordem de peso:

1. **A consulta pública por chave é anônima.** O captcha ali é o controle
   anti-automação do portal. Como a chave de acesso é parcialmente previsível
   (CNPJ + série + número sequencial), um extrator sem captcha não coleta
   apenas as notas do usuário — coleta as de qualquer um. Isso está fora do
   propósito do projeto.
2. **A rota autenticada entrega dado melhor.** O XML obtido via e-CPF traz o
   **GTIN**, que a tela HTML da consulta pública não mostra. É justamente o
   campo que resolve a normalização de produtos.
3. **OCR de captcha é frágil.** Quebra a cada mudança de fonte ou de ruído da
   imagem. Um certificado A1 vale por um ano.

> **Para agentes**: se esta restrição for questionada em uma sessão futura,
> a resposta está aqui. Não reabra o caminho do OCR. A rota suportada é a de
> `07_COLETA_ECPF.md`.
