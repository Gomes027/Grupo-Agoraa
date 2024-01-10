import os
import sys
import time
import keyboard
from time import sleep
import pyautogui as pg
import pygetwindow as gw

def abrir_janela(nome_da_janela):
    """ Função que seleciona a planilha do Excel com o título especificado. """
    janelas = gw.getWindowsWithTitle(nome_da_janela)
    if janelas:
        janela = janelas[0]  # Seleciona a primeira janela correspondente
        if janela.isMinimized:
            janela.restore()  # Restaura a janela se estiver minimizada
        try:
            sleep(1)
            janela.activate()
            sleep(1)
        except Exception as e:
            print(f"Erro ao ativar a janela: {e}")
    else:
        print(f"Janela não encontrada.")
        sys.exit()

def procurar_botão(imagem, clicar):
    start_time = time.time()
    while True:
        localizacao = pg.locateOnScreen(imagem, confidence=0.9)

        if localizacao is not None:
            if clicar == "sim":
                x, y, width, height = localizacao
                centro_x = x + width // 2
                centro_y = y + height // 2
                pg.click(centro_x, centro_y, duration=0.5)
                break
            else:
                break

def procesar_pdfs(pasta):
    # Abre o Chrome
    abrir_janela("Google Chrome"); sleep(1)

    # Lista todos os arquivos na pasta
    arquivos = os.listdir(pasta)

    # Filtra apenas arquivos de imagem (ajuste os formatos conforme necessário)
    extensoes = ['.pdf']
    pdfs = [arq for arq in arquivos if os.path.splitext(arq)[1].lower() in extensoes]

    # Variável para controlar a primeira iteração
    primeira_iteracao = True

    for pdf in pdfs:
        # Extrai o nome do arquivo sem a extensão
        nome_sem_extensao = os.path.splitext(pdf)[0]

        # Divide o nome em elementos separados por espaços
        elementos = nome_sem_extensao.split()

        # Remove os dois últimos elementos se houver suficientes
        if len(elementos) > 2:
            elementos = elementos[:-2]
        else:
            elementos = []

        # Junta os elementos restantes de volta em uma string
        fornecedor = ' '.join(elementos)

        procurar_botão(r"Imgs\pesquisar.png", "sim"); sleep(1)
        
        procurar_botão(r"Imgs\procurar_pedido.png", "não"); sleep(1)

        # Escreve o nome do arquivo
        keyboard.write(fornecedor); sleep(5)
        pg.press("enter"); sleep(6)
        
        # Seleciona a imagem
        pg.doubleClick(1063, 521, duration=0.5); sleep(1)

        # Abre o menu de enviar documentos
        procurar_botão(r"Imgs\+.png", "sim"); sleep(1)
        procurar_botão(r"Imgs\documento.png", "não"); sleep(1)
        pg.press("up", presses=6, interval=0.5); sleep(1)
        pg.press("enter")
        procurar_botão(r"Imgs\enviar_pdf.png", "não"); sleep(1)

        # Seleciona a pasta dos pdfs apenas na primeira iteração
        if primeira_iteracao:
            pg.hotkey("ctrl", "l")
            keyboard.write("F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs")
            pg.press("enter"); sleep(3)
            pg.press("tab", presses=6)
            primeira_iteracao = False
        
        keyboard.write(nome_sem_extensao); sleep(1)
        pg.press("enter")
        procurar_botão(r"Imgs\enviar.png", "sim"); sleep(5)
        
        # Realiza alguma ação com o nome
        print(f"PDF Enviado: {nome_sem_extensao}")

    keyboard.wait("home")
    
    for pdf in pdfs:
        dir_arquivo = os.path.join(pasta, pdf)
        os.remove(dir_arquivo)

procesar_pdfs(r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs")