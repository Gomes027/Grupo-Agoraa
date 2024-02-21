import os
import datetime
import pandas as pd

def verificar_e_atualizar_csv_tresmann():
    # Caminhos dos arquivos
    caminho_csv = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Src\Validades\tables\relatorio_tresmann_filtrado.csv"
    caminho_excel = r"F:\COMPRAS\relatorio_tresmann.xlsx"

    # Verifica se o arquivo CSV existe
    if os.path.exists(caminho_csv):
        # Obtém a data de modificação do arquivo CSV
        data_modificacao = datetime.datetime.fromtimestamp(os.path.getmtime(caminho_csv))
        # Compara com a data atual
        if data_modificacao.date() == datetime.datetime.now().date():
            # Se o arquivo foi modificado hoje, apenas lê o CSV
            df = pd.read_csv(caminho_csv, sep=";")
            return df

    # Se o arquivo CSV não existe ou não foi modificado hoje, atualiza a partir do Excel
    df_tresmann = pd.read_excel(caminho_excel)

    # Definindo as lojas selecionadas
    lojas_selecionadas = ["TRESMANN - SMJ", "TRESMANN - STT", "TRESMANN - VIX"]

    # Faz uma copia do Dataframe original
    df_filtrado_tresmann = df_tresmann.copy()

    # Filtrando o DataFrame com base nos critérios especificados
    df_filtrado_tresmann = df_filtrado_tresmann[
        (df_filtrado_tresmann['FORA DO MIX'] == 'NAO') &
        ~(df_filtrado_tresmann['SETOR'].isin(['HORTIFRUTI', 'MATERIAL CONSUMO'])) &
        (df_filtrado_tresmann['LOJA'].isin(lojas_selecionadas))
    ]

    # Adicionando novas colunas de cálculo ao DataFrame
    df_filtrado_tresmann['MEDIA 12M (30D)'] = df_filtrado_tresmann['VENDA ULTIMOS 12 MESES (QTDE)'] / 12
    df_filtrado_tresmann['VENDA MEDIA'] = (df_filtrado_tresmann['VENDA ULTIMOS 30 DIAS (QTDE)'] + df_filtrado_tresmann['MEDIA 12M (30D)']) / 2

    # Especificando as colunas desejadas e reordenando
    colunas = [
        "CODIGO", "NOME", "LOJA", "NOME_FANTASIA",
        "ESTOQUE ATUAL", "VCTO_MEDIO", "VENDA MEDIA"
    ]
    df_filtrado_tresmann = df_filtrado_tresmann[colunas]

    # Salva o DataFrame processado no arquivo CSV
    df_filtrado_tresmann.to_csv(caminho_csv, index=False, sep=";")

    return df_filtrado_tresmann