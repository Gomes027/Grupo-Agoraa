import os
import subprocess
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta
from Descontos.config import DIR_SIMPLIFICACAO_GERENCIAL, USUARIO, SENHA

def procurar_botão(imagem, clicar):
    while True:
        localizacao = pg.locateOnScreen(imagem, confidence=0.9)

        if localizacao is not None:
            if clicar:
                x, y, width, height = localizacao
                centro_x = x + width // 2
                centro_y = y + height // 2
                pg.click(centro_x, centro_y, duration=0.5)
                break
            else:
                break

def obter_datas_formatadas():
    # Obtem a data atual
    data_atual = datetime.now().date()
    
    # Primeiro dia do mês atual
    primeiro_dia_mes = data_atual.replace(day=1)
    
    # Dia de ontem
    dia_ontem = data_atual - timedelta(days=1)
    
    # Formatando as datas para o formato desejado (ddmmyyyy)
    primeiro_dia_mes_formatado = primeiro_dia_mes.strftime('%d%m%Y')
    dia_ontem_formatado = dia_ontem.strftime('%d%m%Y')
    
    return primeiro_dia_mes_formatado, dia_ontem_formatado

def login_simplificação_gerencial():
    # Inicia o Simplificação e seleciona os menus
    os.system(f'start {DIR_SIMPLIFICACAO_GERENCIAL}')
    procurar_botão(r"Imgs\login_usuario.png", clicar=False)
    pg.typewrite(USUARIO); pg.press("tab")
    pg.typewrite(SENHA); pg. press("enter")

def selecionar_menus():
    procurar_botão(r"Imgs\visoes.png", clicar=True)
    procurar_botão(r"Imgs\dashboards.png", clicar=True)
    procurar_botão(r"Imgs\vendas.png", clicar=True)
    pg.doubleClick(); pg.press("down", presses=3)
    pg.press("enter")

    procurar_botão(r"Imgs\periodo.png", clicar=False)
    primeiro_dia_mes, dia_ontem = obter_datas_formatadas()
    pg.typewrite(primeiro_dia_mes); pg.press("tab")
    pg.typewrite(dia_ontem); pg.press("enter")
    pg.hotkey("alt", "e")

    procurar_botão(r"Imgs\exportar.png", clicar=True)
    procurar_botão(r"Imgs\explorer.png", clicar=False)
    pg.typewrite("descontos"); pg.press("enter"); sleep(5)

def fechar_processo(nome_do_processo):
    try:
        subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
    except subprocess.CalledProcessError as e:
        pass

def extrair_descontos():
    login_simplificação_gerencial()
    selecionar_menus(); sleep(15)
    pg.hotkey("alt", "f4")
    fechar_processo("Plataforma.exe")