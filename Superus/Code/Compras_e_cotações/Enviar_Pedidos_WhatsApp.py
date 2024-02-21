import os
import sys
import keyboard
import pyautogui as pg
import pygetwindow as gw
from time import sleep

def abrir_janela(nome_da_janela):
    """ Seleciona a janela especificada e a ativa. """
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

def procurar_botao(imagem, clicar):
    """ Localiza um botão na tela e clica nele se especificado. """
    while True:
        localizacao = pg.locateOnScreen(imagem, confidence=0.9)
        if localizacao:
            if clicar:
                x, y, width, height = localizacao
                pg.click(x + width // 2, y + height // 2, duration=0.5)
            break
        sleep(1)

def procesar_pdfs(arquivos, tipo_de_pedido):
    """ Processa uma lista de arquivos PDF. """
    arquivos_ordenados = sorted(arquivos)

    abrir_janela("WhatsApp")
    primeira_iteracao = True

    for pdf in arquivos_ordenados:
        nome_sem_extensao = os.path.splitext(pdf)[0]
        elementos = nome_sem_extensao.split()

        fornecedor = ' '.join(elementos[:-2]) if len(elementos) > 2 else ''

        if tipo_de_pedido == "cotação":
            procurar_botao(r"Imgs\grupo_cotacao.png", clicar=True)
            caminho_pdf = "F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"

        elif tipo_de_pedido == "compras":
            procurar_botao(r"Imgs\grupo_compras.png", clicar=True)
            caminho_pdf = "F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"


        procurar_botao(r"Imgs\pesquisar.png", clicar=True)
        procurar_botao(r"Imgs\procurar_pedido.png", clicar=False)

        keyboard.write(fornecedor); sleep(1)
        
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

        # Abre o menu de enviar documentos
        procurar_botao(r"Imgs\+.png", clicar=True); sleep(1)
        procurar_botao(r"Imgs\documento.png", clicar=False); sleep(2)
        pg.click(682, 695, duration=1)

        procurar_botao(r"Imgs\enviar_pdf.png", clicar=False)

        if primeira_iteracao == True:
            keyboard.send("ctrl+l"); sleep(1)
            keyboard.write(caminho_pdf); sleep(1)
            pg.press("enter"); sleep(1)
            pg.press("tab", presses=6); sleep(1)
            primeira_iteracao = False

        keyboard.write(nome_sem_extensao); sleep(1)
        pg.press("enter")

        procurar_botao(r"Imgs\enviar.png", clicar=True); sleep(3)
        print(f"PDF Enviado: {nome_sem_extensao}")

        caminho_completo = os.path.join(caminho_pdf, pdf)
        os.remove(caminho_completo)

    abrir_janela("Superus")

DIR_PDF_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"
DIR_PDF_COTACAO = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"

pdfs_cotacao = {arq for arq in os.listdir(DIR_PDF_COTACAO) if arq.lower().endswith(".pdf")}
pdfs_compras = {arq for arq in os.listdir(DIR_PDF_PEDIDOS) if arq.lower().endswith(".pdf")}

if pdfs_cotacao:
    # Processamento dos PDFs de cotação
    procesar_pdfs(pdfs_cotacao, "cotação")

elif pdfs_compras:
    # Processamento dos PDFs de compras
    procesar_pdfs(pdfs_compras, "compras")  