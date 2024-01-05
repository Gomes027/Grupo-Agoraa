import os
import schedule
import subprocess
from time import sleep
import pyautogui as pg

def atualizar_analise_compra_por_estoque():
    # Caminho para o arquivo Excel que você deseja abrir
    caminho_do_arquivo = r"F:\COMPRAS\Analise compra por estoque v4.11.xlsb"

    # Verifique se o arquivo existe
    if os.path.exists(caminho_do_arquivo):
        # Comando para abrir o Excel
        subprocess.run(['start', 'excel', caminho_do_arquivo], shell=True); sleep(20)
    else:
        print("Arquivo não encontrado.")

    while True:
        if pg.locateOnScreen(r"Imgs\vinculos externos.png", confidence=0.8) is not None:
            print("Imagem Encontrada!")
            pg.press("enter"); sleep(1)
            break

    pg.hotkey("alt", "u", "x", "m", interval=1); sleep(3)
    pg.hotkey("ctrl", "alt", "f5"); sleep(180)
    pg.hotkey("ctrl", "shift", "p")
    print("Arquivo Salvo!")

# Agendar para executar todos os dias
schedule.every().day.at("22:30").do(atualizar_analise_compra_por_estoque)

while True:
    schedule.run_pending()
    sleep(1)