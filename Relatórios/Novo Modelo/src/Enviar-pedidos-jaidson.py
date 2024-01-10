import os
import keyboard
from time import sleep
import pyautogui as pg

DIR_JAIDSON = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Imgs\Jaidson"

def processar_imagens(pasta):
    keyboard.wait("insert")
    
    # Lista todos os arquivos na pasta
    arquivos = os.listdir(pasta)

    # Filtra apenas arquivos de imagem (ajuste os formatos conforme necessário)
    extensoes_imagem = ['.png', '.jpg', '.jpeg']
    imagens = [arq for arq in arquivos if os.path.splitext(arq)[1].lower() in extensoes_imagem]

    for imagem in imagens:
        # Extrai o nome do arquivo sem a extensão
        nome_sem_extensao = os.path.splitext(imagem)[0]

        keyboard.write(nome_sem_extensao)
        sleep(3)
        pg.press("right")
        sleep(2)
    
    keyboard.wait("enter")

    for imagem in imagens:
        caminho_imagem = os.path.join(pasta, imagem)
        os.remove(caminho_imagem)

processar_imagens(DIR_JAIDSON)
