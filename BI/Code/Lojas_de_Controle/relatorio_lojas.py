import os
import locale
import pandas as pd
from datetime import datetime
from Lojas_de_Controle.config import DIR_CONTROLE_SMJ, DIR_CONTROLE_STT, DIR_CONTROLE_VIX, DIR_CONTROLE_MCP, DIR_TROCAS_SMJ, DIR_TROCAS_STT, DIR_TROCAS_VIX, DIR_TROCAS_MCP

# Configuração inicial
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def ler_ultimo_valor_nao_nulo(arquivo, coluna):
    """Lê o último valor não nulo de uma coluna específica em um arquivo Excel."""
    df = pd.read_excel(arquivo, engine='xlrd')
    return df.iloc[:, coluna].dropna().iloc[-1]

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

def relatorio_lojas():
    # Lista para armazenar os DataFrames de cada loja
    controles = []
    trocas = []

    for caminho in [DIR_CONTROLE_SMJ, DIR_CONTROLE_STT, DIR_CONTROLE_VIX, DIR_CONTROLE_MCP, DIR_TROCAS_SMJ, DIR_TROCAS_STT, DIR_TROCAS_VIX, DIR_TROCAS_MCP]:
        registros = int(ler_ultimo_valor_nao_nulo(caminho, 9))
        soma_produto = int(ler_ultimo_valor_nao_nulo(caminho, 23))
        custo = ler_ultimo_valor_nao_nulo(caminho, 30)

        # Extrair o nome do arquivo sem a extensão e separar por espaço
        nome_arquivo, extensao = os.path.splitext(os.path.basename(caminho))
        nome_loja = nome_arquivo.split()[-1]

        if "CONTROLE" in caminho:
            controle = pd.DataFrame({
                'LOJA': [nome_loja],
                'QTDE REGISTROS': [registros],
                'CUSTO': [custo],
                'SOMA PROD.': [soma_produto]
            })
            controles.append(controle)
        else:
            troca = pd.DataFrame({
                'LOJA': [nome_loja],
                'QTDE REGISTROS': [registros],
                'CUSTO': [custo],
                'SOMA PROD.': [soma_produto]
            })
            trocas.append(troca)

        os.remove(caminho)

    # Concatenar os DataFrames
    df_controle = pd.concat(controles, ignore_index=True)
    df_trocas = pd.concat(trocas, ignore_index=True)

    # Adicionar linha de total
    total_controle = pd.DataFrame({
        'LOJA': ['Total:'],
        'QTDE REGISTROS': [""],
        'CUSTO': [df_controle['CUSTO'].sum()],
        'SOMA PROD.': [""]
    })

    total_trocas = pd.DataFrame({
        'LOJA': ['Total:'],
        'QTDE REGISTROS': [""],
        'CUSTO': [df_trocas['CUSTO'].sum()],
        'SOMA PROD.': [""]
    })

    df_controle = pd.concat([df_controle, total_controle], ignore_index=True)
    df_trocas = pd.concat([df_trocas, total_trocas], ignore_index=True)

    # Salvar os DataFrames em um arquivo Excel com sheets diferentes
    nome_arquivo_excel = "F:\BI\LOJAS DE CONTROLE.xlsx"
    with pd.ExcelWriter(nome_arquivo_excel, engine='xlsxwriter') as writer:
        df_controle.to_excel(writer, sheet_name='LOJA CONTROLE', index=False)
        df_trocas.to_excel(writer, sheet_name='LOJA TROCAS', index=False)

        # Chamar a função para ajustar a largura das colunas para cada sheet
        ajustar_largura_colunas(writer, df_controle, 'LOJA CONTROLE')
        ajustar_largura_colunas(writer, df_trocas, 'LOJA TROCAS')

    print(f"Os DataFrames foram salvos no arquivo Excel: {nome_arquivo_excel}")