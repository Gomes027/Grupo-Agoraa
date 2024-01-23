import numpy as np
import pandas as pd
from src.Transferencias.Querys.Query_geral import df_query_geral, din_tr, din_vcto_medio
from src.Transferencias.Querys.Analise_de_compra_por_estoque import df_analise_compra_por_estoque

# parametro_1 = EST SMJ > X VDA
# parametro_2 = EST SMJ > X EMB.
# parametro_3 = EST STT > X VDA
# parametro_4 = EXCEDENTE SMJ
# parametro_5 = PORCENTAGEM EXCEDENTE SMJ
# parametro_6 = NECESSIDADE STT
# parametro_7 = PORCENTAGEM PEDIDO

def gerar_transferencia(loja_que_transfere, loja_que_solicita, parametro_1, parametro_2, parametro_3, parametro_4, parametro_5, parametro_6, parametro_7):
    def estoque_lojas(loja_que_transfere, loja_que_solicita, dataframe):
        # Colunas de dados do analise de compra por estoques 
        colunas_selecionadas = ['NOME',
                                f'EST. ATUAL. {loja_que_transfere}', f'VDA MEDIA 30 D. {loja_que_transfere}', f'VDA 30D. {loja_que_transfere}', 
                                f'MAIOR VENDA {loja_que_transfere}', f'est/vda * {loja_que_transfere}', f'PD. AB. {loja_que_transfere}',
                                f'EST. ATUAL. {loja_que_solicita}', f'VDA MEDIA 30 D. {loja_que_solicita}', f'VDA 30D. {loja_que_solicita}', 
                                f'MAIOR VENDA {loja_que_solicita}', f'est/vda * {loja_que_solicita}', f'PD. AB. {loja_que_solicita}']

        # Filtra as colunas específicas selecionadas, e remove resultados duplicados 
        df_analise_compra_por_estoque_filtrado = df_analise_compra_por_estoque[colunas_selecionadas]
        df_analise_compra_por_estoque_filtrado = df_analise_compra_por_estoque_filtrado.drop_duplicates(subset='NOME', keep='first')

        # Faz o merge dos dados filtrados do analise de compra, com o df_merged
        dataframe = pd.merge(dataframe, df_analise_compra_por_estoque_filtrado, on='NOME', how='left')

        return dataframe

    def logicas_condicionais(loja_que_transfere, loja_que_solicita, dataframe):
        def estoque_excedente(row):
            try:
                if row['TR VIX'] == 1:
                    if row['MAIOR VENDA SMJ'] == 0:
                        result = np.floor(parametro_5 * row[f'EST. ATUAL. {loja_que_transfere}'] / row['EMB_TRANSF']) * row['EMB_TRANSF']
                    else:
                        temp_result = np.floor((row[f'EST. ATUAL. {loja_que_transfere}'] - (parametro_4 * row[f'MAIOR VENDA {loja_que_transfere}'])) / row['EMB_TRANSF']) * row['EMB_TRANSF']
                        result = max(0, temp_result)
                    return result * parametro_7  # Multiplica por 0.5 se 'TR VIX' é 1
                else:
                    if row[f'MAIOR VENDA {loja_que_transfere}'] == 0:
                        result = np.floor(parametro_5 * row[f'EST. ATUAL. {loja_que_transfere}'] / row['EMB_TRANSF']) * row['EMB_TRANSF']
                    else:
                        temp_result = np.floor((row[f'EST. ATUAL. {loja_que_transfere}'] - (parametro_4 * row[f'MAIOR VENDA {loja_que_transfere}'])) / row['EMB_TRANSF']) * row['EMB_TRANSF']
                        result = max(0, temp_result)
                    return result  # Retorna o resultado sem multiplicar por 0.5
            except Exception as e:
                return np.nan  # Retorna NaN em caso de erro
            
        def necessidade_estoque(row):
            try:
                # Verifica se EMB_TRANSF é 0 e retorna 'EXCEDENTE SMJ' se for o caso
                if row[f'MAIOR VENDA {loja_que_solicita}'] == 0:
                    return row['EXCEDENTE SMJ']  # Corresponde ao Z1170 no Excel

                # Calcula o valor e garante que seja no mínimo 0
                valor_calculado = np.ceil(((row[f'MAIOR VENDA {loja_que_solicita}'] * min(row['VAL EM MESES'], parametro_6)) - row[f'EST. ATUAL. {loja_que_solicita}']) / row['EMB_TRANSF']) * row['EMB_TRANSF']
                return max(valor_calculado, 0)
            except:
                return np.nan  # Retorna "" em caso de qualquer outro erro

        dataframe[f'EST {loja_que_transfere} > X VDA'] = dataframe.apply(
            lambda row: "SIM" if not pd.isna(row[f'EST. ATUAL. {loja_que_transfere}']) and not pd.isna(row[f'VDA 30D. {loja_que_transfere}']) and row[f'EST. ATUAL. {loja_que_transfere}'] > (row[f'VDA 30D. {loja_que_transfere}'] * parametro_1) else ("NÃO" if not pd.isna(row[f'EST. ATUAL. {loja_que_transfere}']) and not pd.isna(row[f'VDA 30D. {loja_que_transfere}']) else np.nan),
            axis=1
        )

        dataframe[f'EST {loja_que_transfere} > X EMB.'] = dataframe.apply(
            lambda row: "SIM" if not pd.isna(row[f'EST. ATUAL. {loja_que_transfere}']) and not pd.isna(row['EMB_TRANSF']) and row[f'EST. ATUAL. {loja_que_transfere}'] > (row['EMB_TRANSF'] * parametro_2) else ("NÃO" if not pd.isna(row[f'EST. ATUAL. {loja_que_transfere}']) and not pd.isna(row['EMB_TRANSF']) else np.nan),
            axis=1
        )

        dataframe[f'EST {loja_que_solicita} > X VDA'] = dataframe.apply(
            lambda row: "SIM" if not pd.isna(row[f'EST. ATUAL. {loja_que_solicita}']) and not pd.isna(row[f'MAIOR VENDA {loja_que_solicita}']) and row[f'EST. ATUAL. {loja_que_solicita}'] > (row[f'MAIOR VENDA {loja_que_solicita}'] * parametro_3) else ("NÃO" if not pd.isna(row[f'EST. ATUAL. {loja_que_solicita}']) and not pd.isna(row[f'MAIOR VENDA {loja_que_solicita}']) else np.nan),
            axis=1
        )

        dataframe[f'EXCEDENTE {loja_que_transfere}'] = dataframe.apply(estoque_excedente, axis=1)

        dataframe.loc[:, f'NECESSIDADE {loja_que_solicita}'] = dataframe.apply(necessidade_estoque, axis=1)
        dataframe.loc[:, f'NECESSIDADE {loja_que_solicita}'] = dataframe[f'NECESSIDADE {loja_que_solicita}'].astype('float64')

        # Realizando as conversões
        dataframe.loc[:, f'NECESSIDADE {loja_que_solicita}'] = pd.to_numeric(dataframe[f'NECESSIDADE {loja_que_solicita}'], errors='coerce')
        dataframe.loc[:, f'EXCEDENTE {loja_que_transfere}'] = pd.to_numeric(dataframe[f'EXCEDENTE {loja_que_transfere}'], errors='coerce')

        dataframe.loc[:, 'PEDIDO'] = dataframe.apply(lambda row: row[f'EXCEDENTE {loja_que_transfere}'] if row[f'NECESSIDADE {loja_que_solicita}'] > row[f'EXCEDENTE {loja_que_transfere}'] else row[f'NECESSIDADE {loja_que_solicita}'], axis=1)
        dataframe['PEDIDO'] = np.ceil(dataframe['PEDIDO'].fillna(0)).astype('int64')

        return dataframe

    def filtros(loja_que_transfere, loja_que_solicita, dataframe):
        # Aplicando os filtros
        dataframe = dataframe[
            (dataframe[f"TR {loja_que_transfere}"] == 0) &
            (dataframe[f"TR {loja_que_solicita}"] == 1) &
            (dataframe[f'EST {loja_que_transfere} > X VDA'] == 'SIM') &
            (dataframe[f'EST {loja_que_transfere} > X EMB.'] == 'SIM') &
            (dataframe[f'EST {loja_que_solicita} > X VDA'] == 'NÃO') &
            (dataframe[f'EXCEDENTE {loja_que_transfere}'] > 0) &
            (dataframe[f'NECESSIDADE {loja_que_solicita}'] > 0) &
            (dataframe['PEDIDO'] > 0)
        ]

        return dataframe

    # Remove as linhas duplicatas de df_query_geral
    df_query_geral.drop_duplicates(subset='NOME', keep='first', inplace=True)

    # Extrai os produtos do Analise de Compra por estoque
    df_query_geral_reduzido = df_query_geral[['NOME', 'CODIGO', 'EMB_TRANSF']]
    df_analise_reduzido = df_analise_compra_por_estoque['NOME']
    df_produtos = pd.merge(df_analise_reduzido, df_query_geral_reduzido, on='NOME', how='left')

    # Da merge nos dataframes din_tr e din_vcto_medio, com base nos codigos de df_produtos
    df_merged_din = pd.merge(df_produtos, din_tr, on='CODIGO', how='left')
    df_merged_din = pd.merge(df_merged_din, din_vcto_medio, on='CODIGO', how='left')

    # Aplica uma função lambda que gera a validade em meses
    df_merged_din['VAL EM MESES'] = df_merged_din['VAL_MÉDIA'].apply(lambda x: np.floor(x / 30) * 0.5 if pd.notnull(x) else np.nan)

    # Faz um merge dos setores, grupos e subgrupos com base nos codigos
    df_merged_grupos = pd.merge(df_merged_din, df_query_geral[['CODIGO', 'SETOR', 'GRUPO', 'SUBGRUPO']], on='CODIGO', how='left')

    # Extrai os dados de estoques das lojas
    df_merged_estoques = estoque_lojas(loja_que_transfere, loja_que_solicita, df_merged_grupos)

    # Executa calculos de logicas condicionais
    df_merged_logicas = logicas_condicionais(loja_que_transfere, loja_que_solicita, df_merged_estoques)
    
    # Faz os filtros
    df_filtrado = filtros(loja_que_transfere, loja_que_solicita, df_merged_logicas)
    
    df_final = df_filtrado[["CODIGO", "PEDIDO"]]

    quantidade_produtos = df_final['CODIGO'].count()

    df_final.to_csv(f"{loja_que_transfere} x {loja_que_solicita} BALANCEAMENTO.csv", index=False, header=False, sep=";")

    print(f"Transferência {loja_que_transfere} x {loja_que_solicita}: {quantidade_produtos} Produtos")