import pandas as pd
import matplotlib.pyplot as plt
import os

def carregar_dados(file_path):
    """Carrega os dados do Excel e ajusta a formatação."""
    df = pd.read_excel(file_path)
    df['Data'] = pd.to_datetime(df['Data'], format='%d/%m/%Y', errors='coerce')
    df['Mês'] = df['Data'].dt.to_period('M')
    return df

def analisar_dados(df, output_dir):
    """Executa análises e salva gráficos."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Gasto total por categoria
    gasto_categoria = df.groupby('Categoria')['Total'].sum().sort_values(ascending=False)
    plt.figure(figsize=(10, 6))
    gasto_categoria.plot(kind='bar', title='Gasto Total por Categoria', ylabel='Total Gasto (R$)', xlabel='Categoria')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/gasto_por_categoria.png")
    plt.close()

    # Frequência de compras por produto
    frequencia_compras = df['Descrição'].value_counts().head(10)
    plt.figure(figsize=(10, 6))
    frequencia_compras.plot(kind='bar', title='Top 10 Produtos Mais Comprados', ylabel='Frequência', xlabel='Produto')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/frequencia_compras.png")
    plt.close()

    # Supermercados com maior gasto total
    gasto_supermercado = df.groupby('Supermercado')['Total'].sum().sort_values(ascending=False).head(10)
    plt.figure(figsize=(10, 6))
    gasto_supermercado.plot(kind='bar', title='Top 10 Supermercados com Maior Gasto', ylabel='Total Gasto (R$)', xlabel='Supermercado')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/gasto_supermercado.png")
    plt.close()

    # Evolução mensal do gasto total
    gasto_mensal = df.groupby('Mês')['Total'].sum()
    plt.figure(figsize=(10, 6))
    gasto_mensal.plot(kind='line', marker='o', title='Evolução Mensal do Gasto Total', ylabel='Total Gasto (R$)', xlabel='Mês')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/gasto_mensal.png")
    plt.close()

    # Gasto por Dia da Semana
    df['Dia da Semana'] = df['Data'].dt.day_name(locale='pt_BR.utf8').str.capitalize()
    gasto_por_dia = df.groupby('Dia da Semana')['Total'].sum()
    plt.figure(figsize=(10, 6))
    gasto_por_dia.plot(kind='bar', title='Gasto Total por Dia da Semana', ylabel='Total Gasto (R$)', xlabel='Dia da Semana')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(f"{output_dir}/gasto_por_dia.png")
    plt.close()

    print(f"Gráficos salvos em {output_dir}")

def salvar_relatorio(df, output_dir):
    """Salva um relatório consolidado em Excel."""
    relatorio_path = os.path.join(output_dir, "relatorio_consolidado.xlsx")
    with pd.ExcelWriter(relatorio_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados')
        df.groupby('Categoria')['Total'].sum().to_excel(writer, sheet_name='Gasto por Categoria')
        df.groupby('Supermercado')['Total'].sum().to_excel(writer, sheet_name='Gasto por Supermercado')
    print(f"Relatório salvo em {relatorio_path}")

# Configuração principal
file_path = 'compras.xlsx'
output_dir = 'analises'

# Carregar dados, executar análises e salvar resultados
df = carregar_dados(file_path)
analisar_dados(df, output_dir)
salvar_relatorio(df, output_dir)
