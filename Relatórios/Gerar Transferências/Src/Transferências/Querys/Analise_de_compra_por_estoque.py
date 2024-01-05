import pandas as pd

# Carrega o Dataframe df_analise_compra_por_estoque
df_analise_compra_por_estoque = pd.read_excel(r"F:\COMPRAS\Analise compra por estoque v4.11.xlsb", sheet_name="ANÁLISE_ESTOQUE_LOJA", header=9)
df_analise_compra_por_estoque = df_analise_compra_por_estoque.rename(columns={'Rótulos de Linha': 'NOME'})

# Sufixos e nomes das lojas correspondentes
sufixos_lojas = {'.1': 'STT', 
                 '.2': 'VIX', 
                 '.3': 'MCP', 
                 '': 'SMJ'}  # Sem sufixo para a primeira série de colunas

# Colunas originais a serem renomeadas
colunas_originais = ['EST. ATUAL.', 'VDA MEDIA 30 D.', 'VDA 30D.', 'MAIOR VENDA', 'est/vda *', 'PD. AB.']

# Renomeando as colunas
for sufixo, loja in sufixos_lojas.items():
    for coluna in colunas_originais:
        coluna_nova = f"{coluna}{sufixo}"
        if coluna_nova in df_analise_compra_por_estoque.columns:
            df_analise_compra_por_estoque.rename(columns={coluna_nova: f"{coluna} {loja}"}, inplace=True)