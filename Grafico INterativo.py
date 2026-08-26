import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

# Ler a planilha 'compras.xlsx'
arquivo_excel = 'compras.xlsx'

try:
    # Ler os dados da planilha
    df = pd.read_excel(arquivo_excel, engine='openpyxl')
    print(f"Planilha '{arquivo_excel}' carregada com sucesso!")
except FileNotFoundError:
    print(f"Arquivo '{arquivo_excel}' não encontrado. Verifique o caminho.")
    exit()

# Verificar as primeiras linhas do DataFrame
print("Primeiras linhas do DataFrame:")
print(df.head())

# Converter colunas de data para o formato correto, se necessário
if 'Data' in df.columns:
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')

# Gráfico 1: Frequência de compra por categoria
frequencia_categoria = df['Categoria'].value_counts().reset_index()
frequencia_categoria.columns = ['Categoria', 'Frequência']

fig1 = px.bar(
    frequencia_categoria,
    x='Categoria',
    y='Frequência',
    title='Frequência de Compra por Categoria',
    labels={'Frequência': 'Quantidade'}
)
fig1.show()

# Gráfico 2: Preço mais baixo por produto e supermercado
if 'Descrição' in df.columns and 'Total' in df.columns:
    preco_min = df.groupby(['Descrição', 'Supermercado'])['Total'].min().reset_index()
    fig2 = px.bar(
        preco_min,
        x='Descrição',
        y='Total',
        color='Supermercado',
        title='Preço Mais Baixo por Produto e Supermercado',
        labels={'Total': 'Preço Mais Baixo', 'Descrição': 'Produto'}
    )
    fig2.show()

# Gráfico 3: Quantidade média comprada por mês
if 'Quantidade' in df.columns and 'Data' in df.columns:
    df['Mês'] = df['Data'].dt.to_period('M')
    qtd_media_mes = df.groupby('Mês')['Quantidade'].mean().reset_index()
    fig3 = px.line(
        qtd_media_mes,
        x='Mês',
        y='Quantidade',
        title='Quantidade Média Comprada por Mês',
        labels={'Mês': 'Mês', 'Quantidade': 'Média de Quantidade'}
    )
    fig3.update_traces(mode='lines+markers')
    fig3.show()

# Gráfico 4: Gastos totais por supermercado
if 'Total' in df.columns and 'Supermercado' in df.columns:
    gastos_supermercado = df.groupby('Supermercado')['Total'].sum().reset_index()
    fig4 = px.pie(
        gastos_supermercado,
        names='Supermercado',
        values='Total',
        title='Gastos Totais por Supermercado'
    )
    fig4.show()

# Gráfico 5: Evolução do gasto total ao longo do tempo
if 'Data' in df.columns and 'Total' in df.columns:
    gasto_por_data = df.groupby('Data')['Total'].sum().reset_index()
    fig5 = px.line(
        gasto_por_data,
        x='Data',
        y='Total',
        title='Evolução do Gasto Total ao Longo do Tempo',
        labels={'Data': 'Data', 'Total': 'Gasto Total'}
    )
    fig5.update_traces(mode='lines+markers')
    fig5.show()

print("Gráficos gerados com sucesso!")
