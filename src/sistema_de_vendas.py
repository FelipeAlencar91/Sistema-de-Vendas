import pandas as pd
import os

print('\n=== SISTEMA DE ANÁLISE DE VENDAS ===')

print("Iniciando sistema de vendas...")

# Caminho até o arquivo Excel
caminho_arquivo = os.path.join(
    os.path.dirname(__file__),
    "..",
    "data",
    "vendas.xlsx"
)

print("Caminho do arquivo:", caminho_arquivo)

# Ler o arquivo Excel
df = pd.read_excel(caminho_arquivo)

print("\nDados carregados:")
print(df)

# Criar coluna de total
df["Total"] = df["Quantidade"] * df["Preco"]

print("\nDados com total:")
print(df)

# Faturamento total
faturamento_total = df["Total"].sum()

print("\nFaturamento total:")
print(faturamento_total)

# O produto mais vendido
print('\nProduto_mais_vendido:')
quantidade_mais_vendida= df.groupby('Produto')['Quantidade'].sum().max()
produto_mais_vendido= df.groupby('Produto')['Quantidade'].sum().idxmax()

print(produto_mais_vendido)
print("Quantidade_mais_vendida:", quantidade_mais_vendida)

# O produto menos vendido
print('\nProduto_menos_vendido:')
quantidade_menos_vendido= df.groupby('Produto')['Quantidade'].sum().min()
produto_menos_vendido= df.groupby('Produto')['Quantidade'].sum().idxmin()

print(produto_menos_vendido)
print('Quantidade_menos_vendido:', quantidade_menos_vendido)

# Faturamento por produto
print('\nFaturamento por produto: ')
faturamento_por_produto= (
    df.groupby('Produto')['Total']
   .sum()
    .sort_values(ascending=False)
)
print('Faturamento por produto:', faturamento_por_produto)

# EXPORTAR RELATÓRIO
print('\nExportando relatório de faturamento por produto...')

arquivo_saida = os.path.join(
    os.path.dirname(__file__),
    "..",
    "data",
    "relatorio_faturamento_por_produto.xlsx"
)

faturamento_por_produto.to_excel(arquivo_saida)

print("Relatório gerado com sucesso!")
