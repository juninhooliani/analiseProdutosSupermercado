# 03. Estratégia de Ingestão de Notas Fiscais e Dados

## 1. O que foi apurado sobre as rotas oficiais

Não existe download em massa oficial de XML de NFC-e para **pessoa física**.
As rotas foram verificadas uma a uma:

| Rota | Situação | Motivo |
|---|---|---|
| `NFeDistribuicaoDFe` (lote com certificado) | Indisponível | Cobre apenas modelo 55. NFC-e é modelo 65. |
| `SAE-NFC-e` (lote: 2.000 chaves/consulta, janela de 100 dias) | Indisponível | Nota Técnica v1.0.0: *"o contribuinte a ser pesquisado será o que possui o CNPJ que consta no certificado digital (e-CNPJ)"*. Consulta por CPF não é permitida. |
| Exportação no portal da NFP | Não existe | Nem XML, nem CSV. Somente tela HTML. |
| **Login por e-CPF no portal da NFP** | **Aceito** | O perfil Consumidor aceita `e-CPF ou e-CNPJ, tipo A1 ou A3`. |
| XML por chave, como consumidor identificado | Permitido | Desde que o CPF tenha sido informado na compra. |

**Conclusão**: a melhor rota disponível é a **sessão autenticada por e-CPF** no
portal da NFP. Ver `07_COLETA_ECPF.md` para o detalhamento e o estado da
validação.

## 2. Sobre o reCAPTCHA

A versão anterior deste documento afirmava que a raspagem era *"inviável sem
serviços pagos de quebra de captcha"*. Isso vale para o acesso **anônimo**, mas
não para o **autenticado**: um handshake TLS com certificado de cliente ocorre
abaixo da camada HTTP, e não há captcha em um handshake.

A quebra de captcha por OCR foi **descartada como estratégia do projeto**.
Ver a justificativa em `06_AI_AGENT_RULES.md`, seção 3.

## 3. Métodos de Ingestão do Sistema

### Método 1 (principal): Sessão autenticada por e-CPF
Coleta as chaves de acesso no portal da NFP e baixa o XML de cada nota.
Entrega o dado mais rico disponível — **incluindo o GTIN de cada item**.
Detalhamento em `07_COLETA_ECPF.md`.

### Método 2: Upload manual de XML / HTML / TXT / PDF
O usuário arrasta os arquivos para o sistema. O parser extrai cabeçalho
(data, valor total, CNPJ, estabelecimento) e itens. É o fallback quando a
coleta automatizada não cobre uma nota.

### Método 3: Leitura via QR Code (Mobile / WebCam)
Câmera apontada para o QR Code impresso no cupom. O sistema obtém a URL da
NFC-e e extrai os itens. Serve para captura **no momento da compra** — não
resolve histórico retroativo.

### Método 4: Importação de extratos bancários (OFX / CSV / PDF)
Reconcilia o valor global gasto no cartão com as notas fiscais importadas.

## 4. Pipeline de Normalização de Produtos

A estratégia depende de o item ter ou não GTIN (código de barras).

### 4.1. Com GTIN — caminho preferencial
Disponível quando a origem é **XML** (Métodos 1 e 2). O GTIN é a chave de
identidade do produto: `LEITE UHT INT ELEGANCE 1L` e
`LEITE INTEGRAL ELEGANCE 1000ML` colapsam no mesmo código, entre lojas
diferentes, sem heurística nenhuma.

```text
XML -> GTIN -> tabela `produtos` -> comparação direta de preço entre estabelecimentos
```

### 4.2. Sem GTIN — fallback por texto
Necessário quando a origem é HTML, TXT ou PDF (Métodos 2, 3 e 4), onde o
código de barras normalmente não aparece.

1. **Entrada bruta**: `"LEITE UHT INT ELEGANCE 1L"`
2. **Sanitização**: remoção de caracteres especiais, stop-words e padrões
   numéricos irrelevantes.
3. **Categorização**: cruzamento com tabela de sinônimos/regex (`categorias.py`)
   ou classificação por IA.
4. **Saída normalizada**: `"Leite Integral 1L"` -> Categoria: `Laticínios`.

> **Atenção**: a comparação de preço entre estabelecimentos feita por 4.2 é
> menos confiável que a de 4.1. Descrições variam por loja e por PDV. Sempre
> que houver GTIN disponível para o mesmo item, ele tem precedência sobre o
> nome normalizado.
