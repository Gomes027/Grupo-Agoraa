import subprocess
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta

def fechar_processo(nome_do_processo):
    try:
        subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
    except subprocess.CalledProcessError as e:
        pass

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
    
    # Dia de ontem
    dia_ontem = data_atual - timedelta(days=1)
    
    # Encontrar a última sexta-feira
    dia_da_semana_atual = data_atual.weekday()  # Segunda é 0, Domingo é 6
    # Ajuste para quando hoje for sexta-feira, para pegar a sexta-feira anterior
    if dia_da_semana_atual == 4:  # Se hoje é sexta
        ultima_sexta = data_atual - timedelta(days=7)
    elif dia_da_semana_atual == 5:  # Se hoje é sábado
        ultima_sexta = data_atual - timedelta(days=1)
    elif dia_da_semana_atual == 6:  # Se hoje é domingo
        ultima_sexta = data_atual - timedelta(days=2)
    else:
        ultima_sexta = data_atual - timedelta(days=dia_da_semana_atual + 3)
    
    # Formatando as datas para o formato desejado (ddmmyyyy)
    ultima_sexta_formatado = ultima_sexta.strftime('%d%m%Y')
    dia_ontem_formatado = dia_ontem.strftime('%d%m%Y')
    
    return ultima_sexta_formatado, dia_ontem_formatado

def iniciar_superus():
    """
    Inicia o aplicativo Superus e aguarda até que esteja pronto para uso.
    """
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
    iniciar_superus()
    pg.click(239, 32, duration=0.5)
    pg.click(253, 47, duration=0.5)
    pg.click(13, 38, duration=3)
    pg.click(419, 281, duration=2)
    pg.press("tab", presses=2)
    pg.press("down", presses=3)
    pg.press("tab", presses=2)
    sleep(3)

    ultima_sexta_formatado, dia_ontem_formatado = obter_datas_formatadas()

    pg.write(ultima_sexta_formatado)
    pg.press("tab"); sleep(1)
    pg.write(dia_ontem_formatado)

    pg.click(502, 639, duration=1)
    pg.click(480, 640, duration=1)

    pg.press("up", presses=2)
    pg.press("enter")
    pg.hotkey("alt", "o")

    pg.click(407, 42, duration=5); sleep(3)
    pg.write("ENCARTE-SMJ")
    pg.press("enter"); sleep(3)

    pg.press("f9"); sleep(3)
    pg.press("down")
    pg.press("enter")
    pg.hotkey("alt", "o")

    pg.click(407, 42, duration=5); sleep(3)
    pg.write("ENCARTE-STT")
    pg.press("enter")
    
def extrair_encartes():
    selecionar_menus(); sleep(15)
    fechar_processo("SUPERUS.EXE")