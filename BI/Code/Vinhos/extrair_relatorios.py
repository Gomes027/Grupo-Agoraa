import subprocess
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta

def fechar_processo(nome_do_processo):
    """Fecha um processo pelo nome."""
    try:
        subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
    except subprocess.CalledProcessError as e:
        pass

def obter_datas_formatadas():
    """Obtém datas formatadas para uso no programa."""
    # Obtem a data atual
    data_atual = datetime.now().date()
    
    # Dia anterior com base no dia atual
    dia_ontem = data_atual - timedelta(days=1)
    
    # Dia anterior do ano passado
    dia_ano_passado = data_atual - timedelta(days=365)
    
    # Dia anterior do ano passado, + 1 dia
    dia_ano_passado_menos_um = dia_ano_passado + timedelta(days=-1)
    
    # Primeiro dia do mês
    primeiro_dia_mes = datetime(data_atual.year, data_atual.month, 1).date()
    
    # Primeiro dia do mês do ano passado
    primeiro_dia_mes_ano_passado = datetime(dia_ano_passado.year, dia_ano_passado.month, 1).date()
    
    # Formatando as datas para o formato desejado (ddmmyyyy)
    dia_ontem_formatado = dia_ontem.strftime('%d%m%Y')
    dia_ano_passado_formatado = dia_ano_passado.strftime('%d%m%Y')
    dia_ano_passado_menos_um_formatado = dia_ano_passado_menos_um.strftime('%d%m%Y')
    primeiro_dia_mes_formatado = primeiro_dia_mes.strftime('%d%m%Y')
    primeiro_dia_mes_ano_passado_formatado = primeiro_dia_mes_ano_passado.strftime('%d%m%Y')
    
    return dia_ontem_formatado, dia_ano_passado_formatado, dia_ano_passado_menos_um_formatado, primeiro_dia_mes_formatado, primeiro_dia_mes_ano_passado_formatado

def iniciar_superus():
    """Inicia o aplicativo Superus."""
    subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
    
    while True:
        if pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
            pg.write("848600"); sleep(3)
            pg.press("enter", presses=2, interval=1)
            break
        
    while True:
        if pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
            break

def selecionar_menus():
    """Seleciona menus no aplicativo Superus."""
    # Inicia o superus
    iniciar_superus()

    # Navegação inicial para abrir o menu desejado
    pg.click(242, 33, duration=1)
    pg.click(246, 50, duration=1)
    pg.click(19, 38, duration=3)

    pg.click(194, 496, duration=2)
    pg.press("up", presses=10)
    pg.press("enter")

    pg.click(398, 493, duration=2)
    pg.press("up", presses=10)
    pg.press("enter")

    pg.click(197, 529, duration=1)
    pg.press("up", presses=10)
    pg.press("enter")

    pg.click(501, 635, duration=1)

    resultado = obter_datas_formatadas()

    # Lista de opções a serem selecionadas
    opcoes = [
        "TRESMANN - SMJ",
        "TRESMANN - STT",
        "TRESMANN - VIX",
        "TRESMANN - MCP"
    ]

    pressionamentos = 0
    for opcao in opcoes:
        # Aguarda até que a imagem "aguardar_salvar.png" não seja mais encontrada na tela
        while True:
            pg.doubleClick(514, 357, duration=1)
            pg.write(resultado[0])

            pg.doubleClick(696,357, duration=1)
            pg.write(resultado[0])
            pg.press("enter")

            pg.click(481, 637, duration=1)
            pg.press("up", presses=10)
            pg.press("down", presses=pressionamentos)
            pg.press("enter")
            pg.hotkey("alt", "o"); sleep(3)

            if pg.locateOnScreen(r"Imgs\nenhum registro.png") is not None:
                pg.press("enter"); sleep(3)
            else:
                pg.click(408, 44, duration=5); sleep(3)
                pg.write(f"{opcao} ONTEM"); sleep(2)
                pg.press("enter"); sleep(10)
                pg.press("f9"); sleep(3)

            pg.doubleClick(514, 357, duration=1)
            pg.write(resultado[1])

            pg.doubleClick(696,357, duration=1)
            pg.write(resultado[1])
            pg.hotkey("alt", "o"); sleep(3)

            if pg.locateOnScreen(r"Imgs\nenhum registro.png") is not None:
                pg.press("enter"); sleep(3)
            else:
                pg.click(408, 44, duration=5); sleep(3)
                pg.write(f"{opcao} DIA ANO PASSADO"); sleep(2)
                pg.press("enter"); sleep(10)
                pg.press("f9"); sleep(3)
            
            pg.doubleClick(514, 357, duration=1)
            pg.write(resultado[3])

            pg.doubleClick(696,357, duration=1)
            pg.write(resultado[0])
            pg.hotkey("alt", "o"); sleep(3)

            if pg.locateOnScreen(r"Imgs\nenhum registro.png") is not None:
                pg.press("enter"); sleep(3)
            else:
                pg.click(408, 44, duration=5); sleep(3)
                pg.write(f"{opcao} MES"); sleep(2)
                pg.press("enter"); sleep(10)
                pg.press("f9"); sleep(3)

            pg.doubleClick(514, 357, duration=1)
            pg.write(resultado[4])

            pg.doubleClick(696,357, duration=1)
            pg.write(resultado[2])
            pg.hotkey("alt", "o"); sleep(3)

            if pg.locateOnScreen(r"Imgs\nenhum registro.png") is not None:
                pg.press("enter"); sleep(3)
            else:
                pg.click(408, 44, duration=5); sleep(3)
                pg.write(f"{opcao} MES ANO PASSADO"); sleep(2)
                pg.press("enter"); sleep(10)
                pg.press("f9"); sleep(3)

            pressionamentos += 1

            break
 
def extrair_relatorios():
    """Extrai relatórios usando o Superus."""
    selecionar_menus(); sleep(15)
    fechar_processo("SUPERUS.EXE")