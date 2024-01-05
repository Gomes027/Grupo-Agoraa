import numpy as np
import pandas as pd
from Querys.bases import *
from Querys.Analise_de_compra_por_estoque import df_analise_compra_por_estoque

# Removendo duplicatas
df_query_geral_filtrado.drop_duplicates(subset='NOME', keep='first', inplace=True)

# Realizando o merge dos DataFrames
df_completo = pd.merge(df_analise_compra_por_estoque, df_query_geral_filtrado[['NOME', 'CODIGO', 'EMB_TRANSF']], on='NOME', how='left')

# Convertendo 'CODIGO' e 'EMB_TRANSF' para inteiros, tratando NaNs se necessário
df_completo['CODIGO'] = df_completo['CODIGO'].fillna(0).astype(int)
df_completo['EMB_TRANSF'] = df_completo['EMB_TRANSF'].fillna(0).astype(int)

# Reordenando as colunas
df_merged = df_completo[['NOME', 'CODIGO', 'EMB_TRANSF']]

df_merged = pd.merge(df_merged, din_tr, on='CODIGO', how='left')

# Removendo as colunas 'MERCAPP' e 'Total Geral'
df_merged = df_merged.drop(columns=['MERCAPP', 'Total Geral'])

df_merged = pd.merge(df_merged, din_vcto_medio, on='CODIGO', how='left')

# Supondo que 'df' é o seu DataFrame e 'VCTO_MEDIO' é a coluna equivalente a G5
# Aplicando a operação e tratando erros (assumindo que os erros seriam divisões por zero ou dados não numéricos)
df_merged['VAL EM MESES'] = df_merged['VCTO_MEDIO'].apply(lambda x: np.floor(x / 30) * 0.5 if pd.notnull(x) else "")

df_merged = pd.merge(df_merged, df_query_geral_filtrado[['CODIGO', 'SETOR', 'GRUPO', 'SUBGRUPO']], on='CODIGO', how='left')

# Filtrando as colunas específicas de df_analise_compra_por_estoque
colunas_selecionadas = ['NOME', 'EST. ATUAL. SMJ', 'VDA MEDIA 30 D. SMJ', 'VDA 30D. SMJ', 'MAIOR VENDA SMJ', 
                        'est/vda * SMJ', 'PD. AB. SMJ', 'EST. ATUAL. STT', 'VDA MEDIA 30 D. STT', 'VDA 30D. STT', 
                        'MAIOR VENDA STT', 'est/vda * STT', 'PD. AB. STT']
df_analise_compra_por_estoque_filtrado = df_analise_compra_por_estoque[colunas_selecionadas]

df_analise_compra_por_estoque_filtrado = df_analise_compra_por_estoque_filtrado.drop_duplicates(subset='NOME', keep='first')

# Realizando o merge dos DataFrames com as colunas especificadas
df_merged = pd.merge(df_merged, df_analise_compra_por_estoque_filtrado, on='NOME', how='left')

# Usando o método apply com uma função lambda para aplicar a lógica condicional
df_merged['EST SMJ > X VDA'] = df_merged.apply(lambda row: "SIM" if row['EST. ATUAL. SMJ'] > (row['VDA 30D. SMJ'] * 1) else "NÃO", axis=1)

# Usando o método apply com uma função lambda para aplicar a lógica condicional
df_merged['EST SMJ > X EMB.'] = df_merged.apply(lambda row: "SIM" if row['EST. ATUAL. SMJ'] > (row['EMB_TRANSF'] * 1.5) else "NÃO", axis=1)

df_merged['EST STT > X VDA'] = df_merged.apply(lambda row: "SIM" if row['EST. ATUAL. STT'] > (row['MAIOR VENDA STT'] * 2) else "NÃO", axis=1)

def custom_formula(row):
    try:
        print("Nova Linha:", row)  # Imprime a linha atual

        # Verifica se 'TRESMANN - SMJ' é zero
        if row['TRESMANN - SMJ'] == 0:
            divisor_zero = True
        else:
            divisor_zero = False

        if row['TRESMANN - VIX'] == 1:
            if row['MAIOR VENDA SMJ'] == 0:
                result = 0 if divisor_zero else np.floor(0.20 * row['EST. ATUAL. SMJ'] / row['EMB_TRANSF']) * row['EMB_TRANSF']
            else:
                temp_result = 0 if divisor_zero else np.floor((row['EST. ATUAL. SMJ'] - (1.2 * row['MAIOR VENDA SMJ'])) / row['EMB_TRANSF']) * row['EMB_TRANSF']
                result = max(0, temp_result)
            return result * 0.50
        else:
            print("TRESMANN - VIX não é 1")
            if row['MAIOR VENDA SMJ'] == 0:
                result = np.floor(0.20 * row['EST. ATUAL. SMJ'] / row['EMB_TRANSF']) * row['EMB_TRANSF']
                print("Resultado quando MAIOR VENDA SMJ é 0:", result)
                return result
            else:
                temp_result = np.floor((row['EST. ATUAL. SMJ'] - (1.2 * row['MAIOR VENDA SMJ'])) / row['EMB_TRANSF']) * row['EMB_TRANSF']
                print(temp_result)
                final_result = max(0, temp_result)  # Garante que o resultado não seja negativo
                print("Temp Result:", temp_result, "Final Result:", final_result)
                return final_result if final_result >= 0 else 0  # Retorna 0 se o resultado for negativo
    except Exception as e:
        print(f"Erro: {e}")
        return ""

df_merged['EXCEDENTE SMJ'] = df_merged.apply(custom_formula, axis=1)

df_merged = df_merged.drop_duplicates(subset='NOME', keep='first')