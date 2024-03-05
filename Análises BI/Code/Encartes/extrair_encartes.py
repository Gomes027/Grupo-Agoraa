import subprocess
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta

class GerenciadorDeProcessos:
    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

class InterfaceUsuario:
    @staticmethod
    def procurar_botao(imagem, clicar):
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

    @staticmethod
    def iniciar_superus():
        subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
        while True:
            if pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
                pg.write("848600"); sleep(3)
                pg.press("enter", presses=2, interval=1)
                break
            
        while True:
            if pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
                break

    @staticmethod
    def selecionar_menus(ultima_sexta_formatado, dia_ontem_formatado):
        InterfaceUsuario.iniciar_superus()
        pg.click(239, 32, duration=0.5)
        pg.click(253, 47, duration=0.5)
        pg.click(13, 38, duration=3)
        pg.click(419, 281, duration=2)
        pg.press("tab", presses=2)
        pg.press("down", presses=3)
        pg.press("tab", presses=2)
        sleep(3)

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

class GerenciadorDeDatas:
    @staticmethod
    def obter_datas_formatadas():
        data_atual = datetime.now().date()
        dia_ontem = data_atual - timedelta(days=1)
        dia_da_semana_atual = data_atual.weekday()
        if dia_da_semana_atual >= 4:  # Ajusta para sexta-feira anterior se hoje for sexta, sábado ou domingo
            ajuste = min(dia_da_semana_atual - 4, 2)
            ultima_sexta = data_atual - timedelta(days=ajuste + 1)
        else:
            ultima_sexta = data_atual - timedelta(days=dia_da_semana_atual + 3)
        return ultima_sexta.strftime('%d%m%Y'), dia_ontem.strftime('%d%m%Y')

class ExtratorDeEncartes:
    def __init__(self):
        self.ultima_sexta, self.dia_ontem = GerenciadorDeDatas.obter_datas_formatadas()
    
    def extrair_encartes(self):
        InterfaceUsuario.selecionar_menus(self.ultima_sexta, self.dia_ontem)
        sleep(15)  # Aguarda a finalização das ações anteriores
        GerenciadorDeProcessos.fechar_processo("SUPERUS.EXE")

if __name__ == "__main__":
    extrair = ExtratorDeEncartes()
    extrair.extrair_encartes()