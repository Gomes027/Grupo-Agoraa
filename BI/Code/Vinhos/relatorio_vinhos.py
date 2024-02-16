import os
import locale
import numpy as np
import pandas as pd
from Vinhos.config import * 

# Configuração inicial
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def relatorio_vinhos():
    # Leitura do arquivo Excel
    df = pd.read_excel("F:/COMPRAS/relatorio_tresmann.xlsx", sheet_name="relatorios_tresmann")

    # Filtrar linhas
    df_filtered = df[df["SUBGRUPO"] == "ADEGA"]

    # Selecionar colunas desejadas
    colunas_selecionadas = ["LOJA", "FORA DO MIX", "VENDA ULTIMOS 30 DIAS (QTDE)", "ESTOQUE ATUAL"]
    df_result = df_filtered[colunas_selecionadas]

    return df_result

def ler_ultimo_valor_nao_nulo(arquivo, coluna):
    """Lê o último valor não nulo de uma coluna específica em um arquivo Excel."""
    df = pd.read_excel(arquivo, engine='xlrd')
    return df.iloc[:, coluna].dropna().iloc[-1]

def adicionar_coluna_total(dataframe):
    # Lista de índices desejados
    indices_desejados = ['VENDAS DO DIA ANTERIOR',
                         'ANO ANTERIOR',
                         'VENDAS (MÊS) ACUMULADAS',
                         'ANO  ANTERIOR',
                         'VENDAS ÚLTIMOS 30 DIAS',
                         'ESTOQUE DISPONÍVEL']

    # Adiciona a coluna "TOTAL" ao DataFrame como int
    dataframe['TOTAL'] = 0

    # Preenche os valores na coluna "TOTAL" como a soma das linhas para os índices desejados
    for indice in indices_desejados:
        if indice in dataframe.index:
            total_value = dataframe.loc[indice].sum()
            dataframe.at[indice, 'TOTAL'] = int(total_value) if not pd.isna(total_value) else 0

    # Substitui os valores zero por espaços em branco na coluna "TOTAL"
    dataframe['TOTAL'] = dataframe['TOTAL'].replace(0, '')

    return dataframe


def ajustar_largura_colunas(writer, df, sheet_name):
    """Ajusta a largura das colunas de uma planilha Excel baseada no conteúdo."""
    worksheet = writer.sheets[sheet_name]
    for i, col in enumerate(df.columns):
        col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, col_width)

def obter_nome_loja(caminho):
    # Extrai o último nome da variável do arquivo
    nome_loja = os.path.basename(caminho).replace("DIR_", "").split("_")[-1]
    return nome_loja

def gerar_relatorio_vinhos():
    caminhos_arquivos = [
        ONTEM_SMJ, ONTEM_STT, ONTEM_VIX, ONTEM_MCP,
        MES_SMJ, MES_STT, MES_VIX, MES_MCP,
        ONTEM_ANO_ANTERIOR_SMJ, ONTEM_ANO_ANTERIOR_STT, ONTEM_ANO_ANTERIOR_VIX, ONTEM_ANO_ANTERIOR_MCP,
        MES_ANO_ANTERIOR_SMJ, MES_ANO_ANTERIOR_STT, MES_ANO_ANTERIOR_VIX, MES_ANO_ANTERIOR_MCP
    ]
    
    quantidades = []

    for caminho in caminhos_arquivos:
        try:
            quantidade = int(ler_ultimo_valor_nao_nulo(caminho, 31))
            os.remove(caminho)
        except (FileNotFoundError, ValueError):
            quantidade = 0

        quantidades.append(quantidade)

    # Reformular quantidades para criar o DataFrame corretamente
    quantidades_formatadas = [
        [quantidades[0], quantidades[1], quantidades[2], quantidades[3]],
        [quantidades[8], quantidades[9], quantidades[10], quantidades[11]],
        [quantidades[4], quantidades[5], quantidades[6], quantidades[7]],
        [quantidades[12], quantidades[13], quantidades[14], quantidades[15]]
    ]

    # Nomes das colunas e índices
    colunas = ["SMJ", "STT", "VIX", "MCP"]
    indices = ["VENDAS DO DIA ANTERIOR", "ANO ANTERIOR", "VENDAS (MÊS) ACUMULADAS", "ANO  ANTERIOR"]

    # Criar DataFrame
    df = pd.DataFrame(quantidades_formatadas, columns=colunas, index=indices)

    # Adicionar a linha 'VENDAS ÚLTIMOS 30 DIAS'
    df_vinhos = relatorio_vinhos()
    vendas_ultimos_30_dias = df_vinhos.groupby("LOJA")["VENDA ULTIMOS 30 DIAS (QTDE)"].sum()

    # Adicionar a linha 'ESTOQUE DISPONÍVEL'
    estoque_atual = df_vinhos.groupby("LOJA")["ESTOQUE ATUAL"].sum()

    # Filtrar 'RÓTULOS TOTAIS' por lojas que estão 'FORA DO MIX = NAO'
    rotulos_totais_filtrados = df_vinhos[df_vinhos["FORA DO MIX"] == "NAO"].groupby("LOJA").size()
    
    # Filtrar 'RÓTULOS SEM ESTOQUE' por lojas que estão 'FORA DO MIX = NAO' e 'ESTOQUE ATUAL = 0'
    rotulos_sem_estoque_filtrados = df_vinhos[(df_vinhos["FORA DO MIX"] == "NAO") & (df_vinhos["ESTOQUE ATUAL"] == 0)].groupby("LOJA").size()
    
    # Mapear as lojas para as respectivas colunas no DataFrame
    mapeamento_lojas = {"TRESMANN - SMJ": "SMJ", "TRESMANN - STT": "STT", "TRESMANN - VIX": "VIX", "MERCAPP": "MCP"}
    
    for loja, coluna in mapeamento_lojas.items():
        if loja in vendas_ultimos_30_dias.index:
            df.at["VENDAS ÚLTIMOS 30 DIAS", coluna] = vendas_ultimos_30_dias[loja]
        else:
            df.at["VENDAS ÚLTIMOS 30 DIAS", coluna] = 0

        if loja in estoque_atual.index:
            df.at["ESTOQUE DISPONÍVEL", coluna] = estoque_atual[loja]
        else:
            df.at["ESTOQUE DISPONÍVEL", coluna] = 0
        
        if loja in rotulos_totais_filtrados.index:
            df.at["RÓTULOS TOTAIS", coluna] = int(rotulos_totais_filtrados[loja])
        else:
            df.at["RÓTULOS TOTAIS", coluna] = 0

        if loja in rotulos_sem_estoque_filtrados.index:
            df.at["RÓTULOS SEM ESTOQUE", coluna] = int(rotulos_sem_estoque_filtrados[loja])
        else:
            df.at["RÓTULOS SEM ESTOQUE", coluna] = 0
    
        # Adicionar a linha 'RUPTURA' como a divisão de 'RÓTULOS SEM ESTOQUE' por 'RÓTULOS TOTAIS' em porcentagem
        if df.at["RÓTULOS TOTAIS", coluna] != 0:
            df.at["RUPTURA", coluna] = (df.at['RÓTULOS SEM ESTOQUE', coluna] / df.at['RÓTULOS TOTAIS', coluna]) * 100
        else:
            df.at["RUPTURA", coluna] = 0

    # Arredondando a parte decimal para cima se maior ou igual a 0.5
    df = df.apply(lambda x: x.apply(lambda y: np.ceil(y) if y - int(y) >= 0.5 else y)).astype(int)

    df.iloc[-1] = df.iloc[-1].apply(lambda x: f'{x}%')

    df = adicionar_coluna_total(df)

    # Resetar o índice para transformá-lo em uma coluna normal
    df.reset_index(inplace=True)

    # Renomear a coluna do índice para "DESCRIÇÃO"
    df.rename(columns={'index': 'DESCRIÇÃO'}, inplace=True)

    # Salvar o DataFrame em um arquivo Excel
    nome_arquivo_excel = "F:\BI\VINHOS.xlsx"
    with pd.ExcelWriter(nome_arquivo_excel, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='RESUMO', index=False)

        # Chamar a função para ajustar a largura das colunas para cada sheet
        ajustar_largura_colunas(writer, df, 'RESUMO')

    print(f"O DataFrame foi salvos no arquivo Excel: {nome_arquivo_excel}")