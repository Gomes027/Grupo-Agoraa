import os
import sys
import keyboard
from time import sleep
import pyautogui as pg
import pygetwindow as gw
from Compras_e_cotações.config import *


class AutomacaoWhatsApp:
    def abrir_janela(self, nome_da_janela):
        try:
            janelas = gw.getWindowsWithTitle(nome_da_janela)
            if janelas:
                janela = janelas[0]
                if janela.isMinimized:
                    janela.restore()
                janela.activate()
                sleep(1)
            else:
                print(f"Janela {nome_da_janela} não encontrada.")
                sys.exit()
        except Exception as e:
            print(f"Erro ao ativar a janela: {e}")
            sys.exit()

    def procurar_botao(self, imagem, clicar):
        while True:
            localizacao = pg.locateOnScreen(imagem, confidence=0.9)
            if localizacao:
                if clicar:
                    x, y, width, height = localizacao
                    pg.click(x + width // 2, y + height // 2, duration=0.5)
                break
            sleep(1)

    def processar_pdfs(self, arquivos, tipo_de_pedido):
        arquivos_ordenados = sorted(arquivos)

        self.abrir_janela("WhatsApp")

        for pdf in arquivos_ordenados:
            nome_sem_extensao = os.path.splitext(pdf)[0]
            elementos = nome_sem_extensao.split()
            fornecedor = ' '.join(elementos[:-2]) if len(elementos) > 2 else ''

            if tipo_de_pedido == "cotação":
                self.procurar_botao(r"Imgs\grupo_cotacao.png", clicar=True)
                caminho_pdf = DIR_PDF_COTACAO
            elif tipo_de_pedido == "compras":
                self.procurar_botao(r"Imgs\grupo_compras.png", clicar=True)
                caminho_pdf = DIR_PDF_COMPRAS

            self.interagir_com_fornecedor(fornecedor, caminho_pdf, nome_sem_extensao, pdf)

    def interagir_com_fornecedor(self, fornecedor, caminho_pdf, nome_sem_extensao, pdf):
        self.procurar_botao(r"Imgs\pesquisar.png", clicar=True)
        self.procurar_botao(r"Imgs\procurar_pedido.png", clicar=False)
        keyboard.write(fornecedor)
        sleep(1)

        # Espera até que a mensagem ou o contato seja encontrado, entre outras interações
        self.validar_status_mensagem()

        # Enviar o PDF
        self.enviar_documento(caminho_pdf, nome_sem_extensao)

        print(f"PDF Enviado: {nome_sem_extensao}")
        caminho_completo = os.path.join(caminho_pdf, pdf)
        os.remove(caminho_completo)

    def validar_status_mensagem(self):
        while True:
            if pg.locateOnScreen(r"Imgs\procurando_por_mensagens.png") is not None:
                sleep(5)
            elif pg.locateOnScreen(r"Imgs\nenhuma_msg.png") is not None:
                pg.press("esc")
                break
            elif pg.locateOnScreen(r"Imgs\pedidos_encontrados.png", confidence=0.3) is not None:
                pg.press("enter")
                sleep(3)
                # Seleciona o pedido no WhatsApp
                pg.doubleClick(1063, 521, duration=0.5)
                break

    def enviar_documento(self, caminho_pdf, nome_sem_extensao):
        # Abre o menu de enviar documentos
        self.procurar_botao(r"Imgs\+.png", clicar=True); sleep(1)
        self.procurar_botao(r"Imgs\documento.png", clicar=False); sleep(2)
        pg.click(660, 610, duration=1)

        self.procurar_botao(r"Imgs\enviar_pdf.png", clicar=False)

        keyboard.send("ctrl+l"); sleep(1)
        keyboard.write(caminho_pdf); sleep(1)
        pg.press("enter"); sleep(1)
        pg.press("tab", presses=6); sleep(1)

        keyboard.write(nome_sem_extensao); sleep(1)
        pg.press("enter")

        self.procurar_botao(r"Imgs\enviar.png", clicar=True); sleep(3)