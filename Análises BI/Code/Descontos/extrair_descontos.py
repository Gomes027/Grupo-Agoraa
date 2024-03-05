import os
import subprocess
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta
from config import DIR_SIMPLIFICACAO_GERENCIAL, USUARIO, SENHA

class InterfaceUtils:
    @staticmethod
    def procurar_botão(imagem, clicar):
        while True:
            localizacao = pg.locateOnScreen(imagem, confidence=0.9)
            if localizacao is not None:
                if clicar:
                    x, y, width, height = localizacao
                    centro_x = x + width // 2
                    centro_y = y + height // 2
                    pg.click(centro_x, centro_y, duration=0.1)
                break

class DateHandler:
    @staticmethod
    def obter_datas_formatadas():
        data_atual = datetime.now().date()
        dia_ontem = data_atual - timedelta(days=1)
        return dia_ontem.strftime('%d%m%Y')

class GerenciamentoDescontos:
    @staticmethod
    def login_simplificação_gerencial():
        os.system(f'start {DIR_SIMPLIFICACAO_GERENCIAL}')
        InterfaceUtils.procurar_botão(r"Imgs\login_usuario.png", clicar=False)
        pg.typewrite(USUARIO); pg.press("tab"); pg.typewrite(SENHA); pg.press("enter")

    @staticmethod
    def selecionar_menus():
        InterfaceUtils.procurar_botão(r"Imgs\visoes.png", clicar=True)
        InterfaceUtils.procurar_botão(r"Imgs\dashboards.png", clicar=True)
        InterfaceUtils.procurar_botão(r"Imgs\vendas.png", clicar=True)
        pg.doubleClick(); pg.press("down", presses=3); pg.press("enter")

        InterfaceUtils.procurar_botão(r"Imgs\periodo.png", clicar=False)
        dia_ontem = DateHandler.obter_datas_formatadas()
        pg.typewrite("01012024"); pg.press("tab")
        pg.typewrite(dia_ontem); pg.press("enter"); pg.hotkey("alt", "e")

        InterfaceUtils.procurar_botão(r"Imgs\exportar.png", clicar=True)
        InterfaceUtils.procurar_botão(r"Imgs\explorer.png", clicar=False)
        pg.typewrite("descontos"); pg.press("enter"); sleep(5)

    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

    @staticmethod
    def extrair_descontos():
        GerenciamentoDescontos.login_simplificação_gerencial()
        GerenciamentoDescontos.selecionar_menus(); sleep(30); pg.hotkey("alt", "f4")
        GerenciamentoDescontos.fechar_processo("Plataforma.exe")

if __name__ == "__main__":
    extrair = GerenciamentoDescontos()
    extrair.extrair_descontos()