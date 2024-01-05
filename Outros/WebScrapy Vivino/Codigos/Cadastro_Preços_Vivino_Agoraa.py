import os
import pandas as pd
import pyautogui as pa
from time import sleep
from datetime import datetime, timedelta

data_atual = datetime.now()
primeiro_dia_semana = data_atual - timedelta(days=data_atual.weekday())
quinto_dia_semana = primeiro_dia_semana + timedelta(days=3)
segunda_feira_formatada = primeiro_dia_semana.strftime('%d/%m')
quinta_feira_formatada = quinto_dia_semana.strftime('%d/%m')
data_formatada = primeiro_dia_semana.strftime('%d%m%Y')
fechamento_formatado = quinto_dia_semana.strftime('%d%m%Y')

nome = f'Vivino X Agoraa {segunda_feira_formatada} - {quinta_feira_formatada}'
data = data_formatada
fechamento = fechamento_formatado

os.system("cls")

df = pd.read_excel(r'F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Base de dados_CSV.xlsx')

terceira_coluna = df.iloc[:, 1].tolist()
quinta_coluna = df.iloc[:, 4].tolist()
sexta_coluna = df.iloc[:, 5].tolist()

# Abre a guia do superus
pa.hotkey('win', '3')
sleep(0.5)

#clickar em compras
pa.click(412,31,duration=0.5)

#cliclar em concorrências
pa.click(463,121,duration=0.5)

#incluir uma nova concorrência
pa.click(14,47,duration=0.5)

#digitar VivinosxSuperus
pa.write(nome)
pa.press('enter')

#colocar a data de inicio "sempre segunda-feira"
pa.write(data)
pa.press('enter')

#escolher a loja "VIX"
pa.press('down', presses=3, interval=0.5)
pa.press('enter')

#colocar a data de fechamento "sempre quinta-feira"
pa.write(fechamento)

#salvar 'f6'
pa.press('f6')

#cliclar em concorrentes
pa.click(78,161,duration=0.5)

#inclui um concorrente na concorrência
pa.click(22,950,duration=0.5)

#selecionar fornecedor "vivino"
pa.click(382,452,duration=0.5)
pa.write('vivino')
pa.press('enter')
pa.click(851,647,duration=0.5)
pa.press('enter')
pa.write('vivino')
pa.press('enter', presses=2, interval=0.5)

#inclui um concorrente na concorrência
pa.click(22,950,duration=0.5)

#selecionar fornecedor "tresmann - vix"
pa.click(382,452,duration=0.5)
pa.write('tresmann - vix')
pa.press('enter')
pa.click(851,647,duration=0.5)
pa.press('enter')
pa.write('Agoraa')
pa.press('enter', presses=2, interval=0.5)

#incluir fornecedores na lista e salvar
pa.click(371,202,duration=0.5)
pa.click(373,219,duration=0.5)
pa.click(434,194,duration=0.5)
pa.press('up')
pa.click(434,194,duration=0.5)
pa.click(432,248,duration=0.5)

#clicar em produtos
pa.hotkey('alt', 'p')

#adicionar produtos
pa.click(24,938,duration=0.5)
sleep(1)

for produto in terceira_coluna:
    pa.write(str(produto))
    pa.press('enter')
    pa.hotkey('alt', 'o')
    pa.hotkey('alt', 'n') 
    sleep(0.5)

print("Produtos Cadastrados")
print("="*30)

#fechar cadastro de produtos
pa.press('esc')

#cadastrar preços vivino
pa.click(124,37,duration=0.5)
pa.press('down')
pa.hotkey('alt', 'd')

pa.hotkey('ctrl', 'home')
pa.click(707,112,duration=0.5)

for preço_vivino in sexta_coluna:
    pa.write(str(preço_vivino))
    pa.press('down')

print("Preços Vivino Cadastrados")
print("="*30)

#fechar cadastro de preços vivino
pa.hotkey('alt', 'o')

#cadastrar preços agoraa
pa.press('down', presses=2, interval=0.5 )
pa.hotkey('alt', 'd')
pa.hotkey('ctrl', 'home')
pa.click(707,112,duration=0.5)

for preço_agoraa in quinta_coluna:
    pa.write(str(preço_agoraa))
    pa.press('down')

print("Preços Agoraa Cadastrados")
print("="*30)

#sair do cadastro de preços
pa.hotkey('alt', 'o')
pa.click(1255,11, duration=0.5)
sleep(5)

pa.click(176,42, duration=0.5)
sleep(5)

pa.hotkey('alt', 'o')
sleep(2)

pa.press('enter')