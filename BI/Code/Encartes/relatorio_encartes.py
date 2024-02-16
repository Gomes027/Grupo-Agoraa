import pandas as pd
import locale
import os
from datetime import datetime, timedelta
from Encartes.config import CAMINHO_ENCARTE_SMJ, CAMINHO_ENCARTE_STT, ARQUIVO_EXCEL

# Configuração inicial
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def ler_ultimo_valor_nao_nulo(arquivo, coluna):
    """Lê o último valor não nulo de uma coluna específica em um arquivo Excel."""
    df = pd.read_excel(arquivo, engine='xlrd')
    return df.iloc[:, coluna].dropna().iloc[-1]

def formatar_em_reais(valor):
    """Formata valores monetários em reais."""
    return locale.currency(valor, grouping=True)

def calcular_datas_extremas(considerar_penultima=True):
    """Calcula a última sexta-feira e a próxima quinta-feira."""
    hoje = datetime.now()
    dia_semana = hoje.weekday()
    if considerar_penultima and dia_semana == 4:
        sexta = hoje - timedelta(7)
        quinta = hoje - timedelta(1)
    else:
        sexta = hoje - timedelta(days=(dia_semana + 3) % 7)
        quinta = hoje + timedelta(days=(3 - dia_semana) % 7)
    return sexta.strftime('%d/%m/%Y'), quinta.strftime('%d/%m/%Y')

def atualizar_ou_adicionar_linha_excel(arquivo_excel, df_nova):
    """Atualiza ou adiciona uma linha em um arquivo Excel com base nas datas."""
    try:
        tabela_existente = pd.read_excel(arquivo_excel, engine='openpyxl')
    except FileNotFoundError:
        tabela_existente = pd.DataFrame()

    if 'INICIO' in tabela_existente.columns and 'FIM' in tabela_existente.columns:
        # Encontra os índices das linhas que correspondem às datas de início e fim
        indices = tabela_existente[(tabela_existente['INICIO'] == df_nova['INICIO'].values[0]) & (tabela_existente['FIM'] == df_nova['FIM'].values[0])].index
        
        if len(indices) > 0:
            # Se as datas já existem, atualiza os valores para essas datas
            for col in df_nova.columns:
                tabela_existente.loc[indices, col] = df_nova[col].values[0]
        else:
            # Se as datas não existem, adiciona a nova linha
            tabela_existente = pd.concat([tabela_existente, df_nova], ignore_index=True)
    else:
        tabela_existente = df_nova
    
    with pd.ExcelWriter(arquivo_excel, engine='xlsxwriter') as writer:
        tabela_existente.to_excel(writer, index=False, sheet_name="Resumo")
        ajustar_largura_colunas(writer, tabela_existente)

def ajustar_largura_colunas(writer, df):
    """Ajusta a largura das colunas de uma planilha Excel baseada no conteúdo."""
    worksheet = writer.sheets["Resumo"]
    for i, col in enumerate(df.columns):
        col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        worksheet.set_column(i, i, col_width)

def relatorio_encartes():
    # Lê os últimos valores não nulos das colunas especificadas
    dados = {}
    for caminho in [CAMINHO_ENCARTE_SMJ, CAMINHO_ENCARTE_STT]:
        dados[caminho] = {
            'quantidade': int(ler_ultimo_valor_nao_nulo(caminho, 31)),
            'custo': ler_ultimo_valor_nao_nulo(caminho, 37),
            'lucro_bruto': ler_ultimo_valor_nao_nulo(caminho, 40)
        }

    # Calcula totais e percentuais
    qtde_total = sum(d['quantidade'] for d in dados.values())
    valor_total = sum(d['custo'] for d in dados.values())
    lucro_bruto_total = sum(d['lucro_bruto'] for d in dados.values())
    percentual_lucro = (lucro_bruto_total / valor_total * 100) if valor_total else 0
    lucro_por_produto = lucro_bruto_total / qtde_total if qtde_total else 0

    # Prepara a nova tabela
    inicio, fim = calcular_datas_extremas()
    nova_tabela = pd.DataFrame({
        'INICIO': [inicio], 'FIM': [fim],
        'QTDE TOTAL': [qtde_total],
        'VALOR TOTAL': [formatar_em_reais(valor_total)],
        'LUCRO BRUTO TOTAL': [formatar_em_reais(lucro_bruto_total)],
        '% LUCRO': ["{:.2f}%".format(percentual_lucro)],
        'LUCRO / PROD': [formatar_em_reais(lucro_por_produto)]
    })

    atualizar_ou_adicionar_linha_excel(ARQUIVO_EXCEL, nova_tabela)

    # Limpeza final
    for caminho in [CAMINHO_ENCARTE_SMJ, CAMINHO_ENCARTE_STT]:
        os.remove(caminho)