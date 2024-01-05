import pandas as pd

# Carrega o Dataframe df_relatorio_tresmann
df_relatorio_tresmann = pd.read_excel(r"F:\COMPRAS\relatorio_tresmann.xlsx")


# Selecionando as colunas desejadas do relatório tresmann
colunas_selecionadas_query = ["CODIGO", "NOME", "LOJA", "FORA DO MIX", "NOME_FANTASIA", "SETOR", "GRUPO", "SUBGRUPO", "FLAG5", "VCTO_MEDIO", "EMB_TRANSF"]
df_query_geral = df_relatorio_tresmann[colunas_selecionadas_query]

# Filtrando 'df_query_geral' para "FORA DO MIX" igual a "NAO"
df_query_geral = df_query_geral[df_query_geral["FORA DO MIX"] == "NAO"]

# Selecionando as colunas de Validade de SMJ do relatório tresmann
colunas_selecionadas_validades = ["CODIGO", "LOJA", "VALIDADE"]
validades_smj = df_relatorio_tresmann[colunas_selecionadas_validades]
validades_smj_filtrado = validades_smj[df_relatorio_tresmann["LOJA"] == "TRESMANN - SMJ"]

# Realizando o merge com o df_query_geral usando 'CODIGO' e 'LOJA' como chaves de junção
df_query_geral = pd.merge(df_query_geral, validades_smj_filtrado, on=['CODIGO', 'LOJA'], how='left')

# Convertendo 'CODIGO' e 'EMB_TRANSF' para inteiros
df_query_geral['CODIGO'] = df_query_geral['CODIGO'].astype('Int64')
df_query_geral['EMB_TRANSF'] = df_query_geral['EMB_TRANSF'].astype('Int64')


# Cria o Dataframe din_vcto_medio
din_vcto_medio = df_query_geral.groupby('CODIGO')['VCTO_MEDIO'].max().reset_index() # Agrupa por 'CODIGO' e calcular o maior 'VCTO_MEDIO'
din_vcto_medio['VCTO_MEDIO'] = din_vcto_medio['VCTO_MEDIO'].fillna(0).round().astype('Int64') # Transforma VCTO_MEDIO em inteiro
din_vcto_medio.rename(columns={"VCTO_MEDIO": "VAL_MÉDIA"}, inplace=True) # Renomeia a coluna VCTO_MEDIO para VAL_MÉDIA


# Criando a tabela din_tr, pivotando a tabela por lojas
din_tr = df_query_geral.pivot_table(index='CODIGO', columns='LOJA', values='FLAG5', aggfunc='count', fill_value=0)

# Renomeia as lojas
din_tr.rename(columns={"TRESMANN - SMJ": "TR SMJ"}, inplace=True)
din_tr.rename(columns={"TRESMANN - STT": "TR STT"}, inplace=True)
din_tr.rename(columns={"TRESMANN - VIX": "TR VIX"}, inplace=True)
din_tr = din_tr.drop(columns=['MERCAPP']) # Remove a coluna MERCAPP

# Transforma os TRs em inteiro
din_tr['TR SMJ'] = din_tr['TR SMJ'].astype('Int64')
din_tr['TR STT'] = din_tr['TR STT'].astype('Int64')
din_tr['TR VIX'] = din_tr['TR VIX'].astype('Int64')

# Resetando o índice e renomeando a coluna do índice para 'CODIGO'
din_tr.reset_index(inplace=True)