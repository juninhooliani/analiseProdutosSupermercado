import re
import pandas as pd
from datetime import datetime
import os
import locale
import openai
from dotenv import load_dotenv
from categorias import categorias  # Certifique-se de ter o arquivo categorias.py no mesmo diretório
import subprocess

# Configura o locale para o português do Brasil
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

# Configura a chave da API do OpenAI
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

def conversar_com_chatgpt(pergunta):
    """Tenta obter a categoria de um produto usando o ChatGPT."""
    try:
        resposta = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",  # Use "gpt-4" se você tiver acesso
            messages=[
                {"role": "system", "content": "Você é um especialista em classificar produtos de supermercado em categorias."},
                {"role": "user", "content": pergunta},
            ],
            max_tokens=10,
            temperature=0.5,
        )
        return resposta.choices[0].message["content"].strip()
    except Exception as e:
        print(f"Erro ao se conectar com a API: {e}")
        return None

def identificar_categoria_local(nome_produto):
    """Identifica a categoria localmente usando o arquivo categorias.py."""
    for palavra, categoria in categorias.items():
        if palavra.lower() in nome_produto.lower():
            return categoria
    return "Outros"

def processar_cupom(cupom_texto):
    """Extrai informações do cupom fiscal e retorna uma lista de dados estruturados."""
    linhas = cupom_texto.splitlines()
    
    # Localiza o número do cupom fiscal
    numero_cupom = re.search(r'Extrato Nº: (\d+)', cupom_texto)
    numero_cupom = numero_cupom.group(1) if numero_cupom else "Cupom Desconhecido"

    # Localiza a linha com "Endereço:" e pega as duas linhas anteriores como possíveis nomes
    endereco_idx = next((i for i, linha in enumerate(linhas) if "Endereço:" in linha), None)
    if endereco_idx is not None and endereco_idx >= 1:
        estabelecimento = linhas[endereco_idx - 1].strip()
        if endereco_idx >= 2:
            estabelecimento = f"{linhas[endereco_idx - 2].strip()} {estabelecimento}"
    else:
        estabelecimento = "Supermercado Desconhecido"
    
    # Procura pela data e hora
    data_hora = re.search(r'(\d{2}/\d{2}/\d{4} - \d{2}:\d{2}:\d{2})', cupom_texto)
    if data_hora:
        data_hora = datetime.strptime(data_hora.group(1), '%d/%m/%Y - %H:%M:%S')
    else:
        print("Data e hora não encontradas no cupom.")
        data_hora = None

    produtos = []
    for linha in linhas:
        # Regex para identificar linhas de produtos
        match = re.match(r'^(\d+)\s+(\d+)\s+(.+?)\s+([\d,]+)\s+(\w+)\s+X([\d,]+)\s+\(.+\)\s+([\d,]+)$', linha)
        if match:
            numero = int(match.group(1))  # Número do item
            codigo_barras = match.group(2).strip()  # Código de barras
            descricao = match.group(3).strip()  # Descrição
            quantidade = float(match.group(4).replace(',', '.'))  # Quantidade
            unidade = match.group(5).strip()  # Unidade (un, kg, etc.)
            preco_unitario = float(match.group(6).replace(',', '.'))  # Preço unitário
            valor_total = float(match.group(7).replace(',', '.'))  # Valor total
            
            # Tenta usar o ChatGPT para classificar o produto
            categoria = conversar_com_chatgpt(f"Qual é a categoria do produto: {descricao}?")
            if not categoria:
                # Se falhar, usa a categorização local
                categoria = identificar_categoria_local(descricao)
            
            produtos.append({
                'Número do Cupom': numero_cupom,
                'Data': data_hora.strftime('%d/%m/%Y') if data_hora else None,
                'Hora': data_hora.strftime('%H:%M:%S') if data_hora else None,
                'Dia da Semana': data_hora.strftime('%A').capitalize() if data_hora else None,
                'Supermercado': estabelecimento,
                'Número': numero,
                'Código de Barras': codigo_barras,
                'Descrição': descricao,
                'Quantidade': quantidade,
                'Unidade': unidade,
                'Preço Unitário': preco_unitario,
                'Total': valor_total,
                'Categoria': categoria
            })

    return produtos

def atualizar_planilha(dados, arquivo_excel):
    """Atualiza a planilha Excel com os dados fornecidos."""
    # Limpa a planilha existente
    try:
        df_existente = pd.DataFrame()  # Cria uma planilha vazia
    except Exception:
        df_existente = pd.DataFrame()

    # Converte para DataFrame os dados novos
    df_novo = pd.DataFrame(dados)

    # Concatena os dados novos com os existentes
    df_final = pd.concat([df_existente, df_novo], ignore_index=True)

    # Ordena por data (mais antiga primeiro)
    if 'Data' in df_final.columns:
        df_final['Data'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y', errors='coerce')
        df_final.sort_values(by='Data', inplace=True)
        df_final['Data'] = df_final['Data'].dt.strftime('%d/%m/%Y')

    # Salva a planilha atualizada
    df_final.to_excel(arquivo_excel, index=False, engine='openpyxl')
    print(f"Planilha atualizada com sucesso em: {arquivo_excel}")

    # Abre a planilha automaticamente
    try:
        subprocess.run(["start", arquivo_excel], shell=True)
    except Exception as e:
        print(f"Erro ao tentar abrir a planilha: {e}")

def ler_cupons_da_pasta(pasta):
    """Lê todos os arquivos de texto de uma pasta e processa os cupons."""
    todos_dados = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.txt'):
            caminho_arquivo = os.path.join(pasta, arquivo)
            with open(caminho_arquivo, 'r', encoding='utf-8') as f:
                cupom_texto = f.read()
                dados = processar_cupom(cupom_texto)
                todos_dados.extend(dados)
    return todos_dados

# Caminho da pasta onde os arquivos de texto estão armazenados
pasta_cupons = 'cupons_txt'

# Processa todos os cupons da pasta
dados_cupons = ler_cupons_da_pasta(pasta_cupons)

# Atualiza a planilha Excel
arquivo_excel = 'compras.xlsx'
atualizar_planilha(dados_cupons, arquivo_excel)
