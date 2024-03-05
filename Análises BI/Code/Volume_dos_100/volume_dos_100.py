import pandas as pd
import locale

# Configurando a localidade para pt_BR para formatar os valores monetários
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# Definição de constantes
CAMINHO_ARQUIVO = "F:/COMPRAS/relatorio_tresmann.xlsx"
SHEET_NAME = "relatorios_tresmann"
LOJAS_DESEJADAS = ["TRESMANN - SMJ", "TRESMANN - STT", "TRESMANN - VIX"]

def volumes():
    # Carregando a planilha do Excel
    df = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=SHEET_NAME)

    # Renomeando colunas
    df.rename(columns={
        "Ã¿LTIMO CUSTO (": "CUSTO MEDIO (ULTIMOS 30 DIAS)",
        "TIPO_PROMOCAO": "Tipo_Promocao"
    }, inplace=True)

    # Selecionando colunas de interesse
    colunas_selecionadas = [
        "CODIGO", "NOME", "LOJA", "FORA DO MIX", "NOME_FANTASIA", 
        "SETOR", "VENDA ULTIMOS 30 DIAS (QTDE)", "MEDIA VENDA (ULTIMOS 30 DIAS)"
    ]
    df = df[colunas_selecionadas]

    # Filtrando linhas onde 'FORA DO MIX' é 'NAO' e 'LOJA' está em LOJAS_DESEJADAS
    df = df[(df['FORA DO MIX'] == "NAO") & (df['LOJA'].isin(LOJAS_DESEJADAS))]

    # Adicionando a coluna 'VALOR TOTAL'
    df['VALOR TOTAL'] = df['VENDA ULTIMOS 30 DIAS (QTDE)'] * df['MEDIA VENDA (ULTIMOS 30 DIAS)']

    # Ordenando os dados pela coluna 'VALOR TOTAL' em ordem decrescente
    df.sort_values(by='VALOR TOTAL', ascending=False, inplace=True)

    # Criando a tabela resumo para cada loja
    resumo_lojas = {}
    percentuais = [1, 0.025, 0.05, 0.1]

    for loja in LOJAS_DESEJADAS:
        df_loja = df[df['LOJA'] == loja]
        total_itens = df_loja.shape[0]
        total_venda = df_loja['VALOR TOTAL'].sum()
        
        # Calculando a quantidade de itens para cada percentual
        quantidades = [total_itens] + [int(total_itens * p) for p in percentuais[1:]]
        
        # Construindo a tabela resumo
        tabela_resumo = pd.DataFrame({
            'PORCENTAGEM': ['100%', '2,50%', '5%', '10%'],
            'QTDE ITENS': quantidades,
            'VALOR DA VENDA': [df_loja['VALOR TOTAL'].sum()] + [df_loja['VALOR TOTAL'].iloc[:n].sum() for n in quantidades[1:]]
        })

        # Calculando a coluna '% / TOTAL'
        tabela_resumo['% / TOTAL'] = ["100%"] + [(valor / total_venda * 100).round(0).astype(int).astype(str) + '%' for valor in tabela_resumo['VALOR DA VENDA'][1:]]

        # Formatando a coluna 'VALOR DA VENDA' para o formato de moeda
        tabela_resumo['VALOR DA VENDA'] = tabela_resumo['VALOR DA VENDA'].apply(lambda x: locale.currency(x, grouping=True))

        # Armazenando a tabela resumo no dicionário
        resumo_lojas[loja] = tabela_resumo

    # Caminho do arquivo Excel onde os resumos serão salvos
    caminho_arquivo_excel = 'F:/BI/VOLUME DOS 100.xlsx'

    with pd.ExcelWriter(caminho_arquivo_excel, engine='xlsxwriter') as writer:
        for loja, tabela in resumo_lojas.items():
            # Cada loja terá sua própria planilha, nomeada com o nome da loja
            sheet_name = loja[:31]  # Limita o nome da planilha a 31 caracteres, que é o máximo permitido pelo Excel
            tabela.to_excel(writer, sheet_name=sheet_name, index=False)
            
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Ajusta a largura das colunas automaticamente
            for col_num, value in enumerate(tabela.columns.values):
                column_len = tabela[value].astype(str).str.len().max()
                column_len = max(column_len, len(value)) + 2  # Adiciona um pouco mais de espaço
                worksheet.set_column(col_num, col_num, column_len)