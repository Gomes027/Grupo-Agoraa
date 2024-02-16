import os
import datetime
import pandas as pd

def verificar_e_atualizar_csv_precos():
    # Caminhos dos arquivos
    caminho_csv = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Src\Validades\tables\precos_4_lojas_filtrado.csv"
    caminho_excel = r"F:\BI\Bases\relatorio_precos_4_lojas.xlsx"

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
    df_precos = pd.read_excel(caminho_excel)

    # Faz uma copia do Dataframe original
    df_precos_filtrado = df_precos.copy()

    # Definindo as lojas selecionadas
    lojas_selecionadas = ["TRESMANN - SMJ", "TRESMANN - STT", "TRESMANN - VIX"]

    # Filtrando 'df_tresmann' para incluir apenas as lojas selecionadas
    df_precos_filtrado = df_precos_filtrado[df_precos_filtrado['NOME_LOJA'].isin(lojas_selecionadas)]

    # Selecionando colunas específicas de 'df_precos'
    colunas_selecionadas = ["CODIGO", "NOME", "NOME_LOJA", "CUSTOLIQUIDO"]
    df_precos_selecionado = df_precos_filtrado[colunas_selecionadas]

    # Ordenando o DataFrame 'df_precos_selecionado' por "CODIGO"
    df_ordenado = df_precos_selecionado.sort_values(by="CODIGO", ascending=True)

    # Salva o DataFrame processado no arquivo CSV
    df_ordenado.to_csv(caminho_csv, index=False, sep=";")

    return df_ordenado