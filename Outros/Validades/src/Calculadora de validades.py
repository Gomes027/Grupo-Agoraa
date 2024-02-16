import os
import pandas as pd
from datetime import datetime
from prettytable import PrettyTable
from querry.precos_4_lojas import verificar_e_atualizar_csv_precos
from querry.relatorio_tresmann import verificar_e_atualizar_csv_tresmann

# Função para buscar informações do produto
def buscar_info_produto(df, codigo, loja, coluna):
    result = df.loc[(df['CODIGO'] == codigo) & (df['LOJA'] == loja), coluna]
    return None if result.empty else result.iloc[0]

# Funções específicas baseadas na função genérica acima
def buscar_nome_produto(codigo, loja):
    return buscar_info_produto(df_final_tresmann, codigo, loja, 'NOME')

def buscar_vcto_produto(codigo, loja):
    return buscar_info_produto(df_final_tresmann, codigo, loja, 'VCTO_MEDIO')

def buscar_estoque_produto(codigo, loja):
    return buscar_info_produto(df_final_tresmann, codigo, loja, 'ESTOQUE ATUAL')

def buscar_venda_media_produto(codigo, loja):
    return buscar_info_produto(df_final_tresmann, codigo, loja, 'VENDA MEDIA')

df_final_tresmann = verificar_e_atualizar_csv_tresmann()
df_final_precos = verificar_e_atualizar_csv_precos()

# Trabalhando com datas
data_validade = '06/02/2024'
data_validade = pd.to_datetime(data_validade, format='%d/%m/%Y')
data_hoje = pd.to_datetime(datetime.now().strftime('%d/%m/%Y'), format='%d/%m/%Y')
dias_a_vencer = (data_validade - data_hoje).days

# Cálculos de projeção e estoque
codigo, loja, quantidade = 44860, "TRESMANN - VIX", 20
venda_media_30_dias = buscar_venda_media_produto(codigo, loja)
projecao_de_venda = ((venda_media_30_dias / 30) * dias_a_vencer) - 2
estoque_atual = buscar_estoque_produto(codigo, loja)
primeira_condicao = (estoque_atual + quantidade) if projecao_de_venda < 0 else (estoque_atual + quantidade - projecao_de_venda)
resultado = 0 if primeira_condicao < 0 else primeira_condicao

# Cálculo do custo líquido
def buscar_custoliquido_produto(codigo, loja, resultado):
    result = df_final_precos.loc[(df_final_precos['CODIGO'] == codigo) & (df_final_precos['NOME_LOJA'] == loja), 'CUSTOLIQUIDO']
    if result.empty:
        return None
    custo_liquido = result.iloc[0] * resultado
    return f'R${custo_liquido:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

perda_financeira_projetada = buscar_custoliquido_produto(codigo, loja, resultado)

# Lógica condicional final
vcto_medio = buscar_vcto_produto(codigo, loja)
if projecao_de_venda > (quantidade + estoque_atual):
    resultado = "RECEBER"
elif dias_a_vencer > (vcto_medio - (vcto_medio * 0.10)):
    resultado = "RECEBER"
else:
    resultado = "RECUSAR"

nome_produto = buscar_nome_produto(codigo, loja)

os.system("cls")

# Criando a tabela
tabela = PrettyTable()

# Adicionando as colunas
tabela.field_names = ["Informação", "Valor"]

# Adicionando as linhas
tabela.add_row(["Código", codigo])
tabela.add_row(["Produto", nome_produto])
tabela.add_row(["Dias a Vencer", dias_a_vencer])
tabela.add_row(["Vcto Médio", vcto_medio])
tabela.add_row(["Estoque Atual", estoque_atual])
tabela.add_row(["Venda Média 30 Dias", venda_media_30_dias])
tabela.add_row(["Projeção de Venda", projecao_de_venda])
tabela.add_row(["Perda Financeira Projetada", perda_financeira_projetada])
tabela.add_row(["Resultado", resultado])

# Exibindo a tabela
print(tabela)