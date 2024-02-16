import locale
import pandas as pd
from Descontos.config import CAMINHO_ARQUIVO_ENTRADA, CAMINHO_ARQUIVO_SAIDA

def configurar_locale_para_brasil():
    """Configura o locale para Português do Brasil."""
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def formatar_como_moeda(valor):
    """Formata um valor como moeda ou retorna NaN se o valor não for numérico."""
    if pd.isna(valor):
        return valor
    else:
        return locale.currency(valor, grouping=True)

def carregar_dados_do_excel(caminho):
    """Carrega dados de um arquivo Excel para um DataFrame do Pandas."""
    return pd.read_excel(caminho)

def converter_tipos_de_colunas(df):
    """Converte os tipos das colunas do DataFrame."""
    df['Loja'] = df['Loja'].astype('Int64')
    df['Data Hora'] = pd.to_datetime(df['Data Hora'], format="%d/%m/%Y %H:%M:%S", errors='coerce')
    df['PDV'] = df['PDV'].astype('Int64')
    df['Cupom'] = df['Cupom'].astype('Int64')
    df['Produto*'] = df['Produto*'].astype('str')
    df['Qtde*'] = df['Qtde*'].astype('float')
    df['Total s/ Desconto'] = df['Total s/ Desconto'].astype('float')
    df['Desconto'] = pd.to_numeric(df['Desconto'].replace('[\$,]', '', regex=True), errors='coerce')
    df['Total c/ Desconto'] = df['Total c/ Desconto'].astype('float')
    df['Usuário Desconto'] = df['Usuário Desconto'].astype('str')
    df['% Desconto'] = df['% Desconto'].astype('float')
    df['Motivo'] = df['Motivo'].astype('str')
    return df

def remover_ultima_linha(df):
    """Remove a última linha do DataFrame."""
    return df.iloc[:-1]

def criar_tabela_pivotada(df):
    """Cria uma tabela pivotada com os dados tratados."""
    pivot = df.pivot_table(values='Desconto', index='Motivo', columns='Loja', 
                           aggfunc='sum', margins=True, margins_name='Total Geral')
    pivot = pivot.rename_axis(index='DESCONTOS')  # Define explicitamente o nome do índice
    pivot.rename(columns={1: 'SMJ', 2: 'STT', 3: 'VIX'}, inplace=True)
    return pivot

def formatar_tabela_pivotada_com_moeda(pivot):
    """Aplica a formatação de moeda para todas as células da tabela pivotada."""
    return pivot.map(formatar_como_moeda)

def salvar_tabela_pivotada(pivot, caminho_saida):
    """Salva a tabela pivotada formatada em um arquivo Excel."""
    pivot.to_excel(caminho_saida)

# Fluxo principal do programa
def tabela_dinamica():
    configurar_locale_para_brasil()
    df = carregar_dados_do_excel(CAMINHO_ARQUIVO_ENTRADA)
    df = converter_tipos_de_colunas(df)
    df = remover_ultima_linha(df)
    pivot_table = criar_tabela_pivotada(df)
    pivot_table_formatada = formatar_tabela_pivotada_com_moeda(pivot_table)
    salvar_tabela_pivotada(pivot_table_formatada, CAMINHO_ARQUIVO_SAIDA)