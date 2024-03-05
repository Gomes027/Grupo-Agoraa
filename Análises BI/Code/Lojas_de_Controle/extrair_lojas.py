import subprocess
import pyautogui as pg
from time import sleep

class ProcessManager:
    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

    @staticmethod
    def iniciar_processo(caminho_executavel=r"C:\SUPERUS VIX\Superus.exe"):
        subprocess.Popen([caminho_executavel])

class SuperusApp:
    @staticmethod
    def iniciar():
        ProcessManager.iniciar_processo()
        while True:
            if pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
                pg.write("848600"); sleep(3)
                pg.press("enter", presses=2, interval=1)
                break
            
        while True:
            if pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
                break

    @staticmethod
    def selecionar_menus():    
        # Navegação inicial para abrir o menu desejado
        pg.click(382, 37, duration=1)
        pg.click(380, 54, duration=1)
        pg.click(282, 32, duration=3)
        pg.click(216, 148, duration=2)
        pg.click(1087, 40, duration=1)
        pg.click(1096, 338, duration=1)
        pg.click(665, 710, duration=1)
        sleep(3)

        # Seleciona a opção
        pg.press("down", presses=10); sleep(2)
        pg.press("enter")
        pg.hotkey("alt", "o")
        pg.click(409, 43, duration=5)
        sleep(5)

        # Lista de opções a serem selecionadas
        opcoes = [
            "CONTROLE INTERNO - MCP",
            "CONTROLE INTERNO - VIX",
            "CONTROLE INTERNO - STT",
            "CONTROLE INTERNO - SMJ",
            "TROCAS - MCP",
            "TROCAS - STT",
            "TROCAS - SMJ",
            "TROCAS - VIX"
        ]

        pressionamentos_up = 1  # Inicia com 2 pressionamentos

        for i, opcao in enumerate(opcoes, start=1):  # Inicia a contagem em 1
            pg.write(opcao); sleep(2)
            pg.press("enter"); sleep(10)

            # Aguarda até que a imagem "aguardar_salvar.png" não seja mais encontrada na tela
            while True:
                if pg.locateOnScreen(r"Imgs\aguardar_salvar.png", confidence=0.9) is None:
                    break

            pg.press("f9"); sleep(3)

            # Na quarta iteração, lista os negativos
            if i == 4:
                pg.click(1096, 338, duration=1)
                pg.click(665, 710, duration=1)
                sleep(3)
                
            pg.press("down", presses=10)
            pg.press("up", pressionamentos_up)
            pg.press("enter"); sleep(10)
            pressionamentos_up += 1  # Incrementa para a próxima iteração

            pg.hotkey("alt", "o")
            pg.click(409, 43, duration=8)
            sleep(5)
        sleep(15)

class Main:
    def executar(self):
        SuperusApp.iniciar()
        SuperusApp.selecionar_menus()
        ProcessManager.fechar_processo("SUPERUS.EXE")

if __name__ == "__main__":
    app = Main()
    app.executar()