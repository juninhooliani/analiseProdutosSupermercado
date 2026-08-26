import re
import pandas as pd
from datetime import datetime
from categorias import categorias

def processar_cupom(cupom_texto):
    """Extrai informações do cupom fiscal e retorna uma lista de dados estruturados."""
    linhas = cupom_texto.splitlines()
    
    supermercado = re.search(r'^(.*) LTDA', linhas[0])
    supermercado = supermercado.group(1).strip() if supermercado else "Supermercado Desconhecido"
    
    # Procura pela data e hora
    data_hora = re.search(r'(\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2})', cupom_texto)
    if data_hora:
        data_hora = datetime.strptime(data_hora.group(1), '%d/%m/%Y - %H:%M:%S')
    else:
        data_hora = None

    produtos = []
    for linha in linhas:
        # Identifica linhas de produtos pelo padrão (CÓD DESCRIÇÃO QTD UN VL UN VL ITEM)
        match = re.match(r'\d+\s+\d+\s+([\w\s]+)\s+([\d,]+)\s+\w+\s+X([\d,]+)\s+\([\d,]+\)\s+([\d,]+)', linha)
        if match:
            descricao = match.group(1).strip()
            quantidade = float(match.group(2).replace(',', '.'))
            preco_unitario = float(match.group(3).replace(',', '.'))
            valor_total = float(match.group(4).replace(',', '.'))
            
            produtos.append({
                'Data': data_hora.date() if data_hora else None,
                'Hora': data_hora.time() if data_hora else None,
                'Dia da Semana': data_hora.strftime('%A') if data_hora else None,
                'Supermercado': supermercado,
                'Descrição do Produto': descricao,
                'Quantidade': quantidade,
                'Preço Unitário': preco_unitario,
                'Valor Total': valor_total,
                # Categoria será processada posteriormente
                'Categoria': identificar_categoria(descricao)
            })

    return produtos

def identificar_categoria(nome_produto):
    """Identifica a categoria com base no nome do produto."""
    for palavra, categoria in categorias.items():
        if palavra.lower() in nome_produto.lower():
            return categoria
    return 'Outros'


def atualizar_planilha(dados, arquivo_excel):
    """Atualiza a planilha Excel com os dados fornecidos."""
    try:
        df_existente = pd.read_excel(arquivo_excel, engine='openpyxl')
    except FileNotFoundError:
        df_existente = pd.DataFrame()

    df_novo = pd.DataFrame(dados)
    df_final = pd.concat([df_existente, df_novo]).drop_duplicates()
    
    # Ordena por data
    df_final.sort_values(by='Data', inplace=True)
    df_final.to_excel(arquivo_excel, index=False, engine='openpyxl')

# Texto do cupom fiscal (substituir pela leitura real do arquivo)
cupom_texto = """
MINIMERCADO INOCOOP LTDA
Endereço: AVENIDA CAETANO DECARO, Nº 1455 - Nao Informado
Bairro: PARQUE RESIDENCIAL LARANJEIRAS I -  CEP: 15904-023 -  TAQUARITINGA - SP
 CNPJ:  48.965.532/0001-90  I.E.:  684156962119  I.M.: 
--------------------------------------------------------------------------------------------------------------
Extrato Nº: 025712
CUPOM FISCAL ELETRÔNICO - SAT
--------------------------------------------------------------------------------------------------------------
CPF/CNPJ do Consumidor: 000.000.000-00
Razão Social/ Nome: XXX
--------------------------------------------------------------------------------------------------------------
#	COD	DESCRIÇÃO	QTD	UN	VL UN R$	(VL TR R$)*	VL ITEM R$
1	015219	ABACAXI	1,0000	UN	X11,99	(0,00)	11,99

2	023679	ACQUAMIX SODA 1,5L COM CASCO	1,0000	UN	X24,50	(0,00)	24,50

3	026917	ACHOCOLATADO MOCOQUINHA 200ML	3,0000	UN	X1,69	(0,00)	5,07

4	026744	SALGADINHO CHEETOS 37G	1,0000	UN	X3,49	(0,00)	3,49

5	023781	CHEETOS 75G	1,0000	UN	X6,99	(0,00)	6,99

6	026744	SALGADINHO CHEETOS 37G	1,0000	UN	X3,49	(0,00)	3,49

7	024159	XAROPE DILUTE 500ML	1,0000	UN	X29,98	(0,00)	29,98

8	026917	ACHOCOLATADO MOCOQUINHA 200ML	1,0000	UN	X1,69	(0,00)	1,69


Total de descontos/ acréscimos sobre o item
0,00
TOTAL R$
87,20

Cartão de Crédito87,20
Troco R$:0,00
Comete crime quem sonega
--------------------------------------------------------------------------------------------------------------
OBSERVAÇÕES DO CONTRIBUINTE

*Valor aproximado dos tributos do item

Valor aproximado dos tributos deste cupom R$
31,39

(conforme Lei Fed. 12.741/2012)
--------------------------------------------------------------------------------------------------------------
SAT Nº 001342635-45
30/12/2023 - 13:17:27
3523 1248 9655 3200 0190 5900 1342 6350 2571 2035 7511
"""

# Processa o cupom fiscal
dados = processar_cupom(cupom_texto)

# Atualiza a planilha
arquivo_excel = 'compras.xlsx'
atualizar_planilha(dados, arquivo_excel)
print("Planilha atualizada com sucesso!")
