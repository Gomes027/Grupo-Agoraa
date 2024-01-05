import os
import sys
import csv
import time
import shutil
import chardet
import datetime
from time import sleep
import pyautogui as pg

def procurar_botão(imagem, clicar):
    start_time = time.time()
    while True:
        localizacao = pg.locateOnScreen(imagem)

        if localizacao is not None:
            if clicar == "sim":
                x, y, width, height = localizacao
                centro_x = x + width // 2
                centro_y = y + height // 2
                pg.click(centro_x, centro_y, duration=0.5)
                break
            else:
                break

def login_simplificação(caminho_simplificação_exe):
    # Inicia o Simplificação e seleciona os menus
    os.system(f'start {caminho_simplificação_exe}')
    procurar_botão(r"Imgs\logar.png", "sim")
    procurar_botão(r"Imgs\operacoes.png", "sim")
    procurar_botão(r"Imgs\loja_loja.png", "sim")

def selecionar_lojas(nome_arquivo):
    # Seleciona as lojas
    partes = nome_arquivo.split()
    
    if len(partes) >= 2:
        loja_1 = partes[0]
        loja_2 = partes[2]

    loja_que_pede = f"Imgs\\{loja_2}.png"
    loja_que_responde = f"Imgs\\{loja_2}\\{loja_1}.png"

    procurar_botão(loja_que_pede, "sim")
    procurar_botão(loja_que_responde, "sim")
    procurar_botão(r"Imgs\digitar.png", "sim")

def processar_arquivo(caminho_arquivo):
    # Digita o pedido
    with open(caminho_arquivo, 'r') as arquivo:
        linhas = arquivo.readlines()
    
    for linha in linhas[:]:
        codigo, quantidade = linha.strip().split(";")
        print(f"{codigo} - {quantidade}", end="\n")
        
        pg.typewrite(codigo); pg.press("enter")
        procurar_botão(r"Imgs\qtde_2.png", "não")

        pg.typewrite(quantidade); pg.press("enter")
        procurar_botão(r"Imgs\menu_digitacao.png", "não"); sleep(1)
            
        linhas.remove(linha)

        with open(caminho_arquivo, 'w') as arquivo:
            arquivo.writelines(linhas)

    procurar_botão(r"Imgs\enviar_arquivo.png", "sim")
    procurar_botão(r"Imgs\confirmar.png", "não"); pg.press("enter")
    procurar_botão(r"Imgs\ok.png", "não"); pg.press("enter")

    sleep(3); pg.hotkey("alt", "f4")
    print("Digitado com sucesso!"); print("\n")

def digitar_transferencia(novo_arquivo, arquivo_completo):
    nome_arquivo_sem_extensao = os.path.splitext(novo_arquivo)[0]
    print(f"Transferência: {nome_arquivo_sem_extensao}")

    # Contar e imprimir o número de linhas do arquivo
    try:
        with open(arquivo_completo, 'r') as arquivo:
            num_linhas = sum(1 for linha in arquivo)
        print(f"Produtos: {num_linhas} Itens")
    except Exception as e:
        print(f"Erro ao ler o arquivo {novo_arquivo}: {e}")

    login_simplificação(r"C:\SimplificAcao_Valid\SmartApp.exe")
    selecionar_lojas(nome_arquivo_sem_extensao)
    processar_arquivo(arquivo_completo)