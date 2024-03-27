import os
import io
import sys
import keyboard
import win32clipboard
from PIL import Image
from time import sleep
import pyautogui as pg
import pygetwindow as gw


class EnviarPedidos:
    def __init__(self, diretorio):
        self.diretorio = diretorio

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
                print(f"Janela '{nome_da_janela}' não encontrada.")
                sys.exit()
        except Exception as e:
            print(f"Erro ao ativar a janela: {e}")
            sys.exit()

    def procurar_botao(self, caminho_imagem, clicar):
        while True:
            try:
                localizacao = pg.locateOnScreen(caminho_imagem, confidence=0.8)
                if localizacao:
                    if clicar:
                        x, y, largura, altura = localizacao
                        pg.click(x + largura // 2, y + altura // 2, duration=0.5)
                    break
            except Exception:
                pass
            
            sleep(1)

    def copiar_imagem_para_area_transferencia(self, caminho_imagem):
        imagem = Image.open(caminho_imagem)
        saida = io.BytesIO()
        imagem.convert("RGB").save(saida, "BMP")
        dados = saida.getvalue()[14:]
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, dados)
        win32clipboard.CloseClipboard()

    def processar_imagens(self, comprador):
        self.abrir_janela("Whatsapp")
        IMGS = os.path.join(self.diretorio, comprador)
        arquivos_de_imagens = os.listdir(IMGS)
        extensoes_imagem = ['.png', '.jpg', '.jpeg']
        imagens = [arquivo for arquivo in arquivos_de_imagens if os.path.splitext(arquivo)[1].lower() in extensoes_imagem]
        self.procurar_botao(r"Imgs\grupo_jaidson.png", True)
        self.procurar_botao(r"Imgs\digite.png", True)

        for imagem in imagens:
            nome_sem_extensao = os.path.splitext(imagem)[0]
            caminho_completo = os.path.join(IMGS, imagem)
            print(caminho_completo)
            self.copiar_imagem_para_area_transferencia(caminho_completo)
            pg.hotkey("ctrl", "v"); sleep(3)
            keyboard.write(nome_sem_extensao); sleep(2)

        #sleep(5); pg.press("enter")

        #for imagem in imagens:
            #caminho_completo = os.path.join(IMGS, imagem)
            #os.remove(caminho_completo)

# Exemplo de uso
if __name__ == "__main__":
    DIR_JAIDSON = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Imgs\automacao.compras1"
    automacao = EnviarPedidos(DIR_JAIDSON)
    automacao.processar_imagens("JAIDSON")