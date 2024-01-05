import os
import sys
import shutil
import schedule
import keyboard
import pandas as pd
from time import sleep
import pyautogui as pg
from datetime import date

# Desabilita o Failsafe do Pyautogui
pg.FAILSAFE = False

def atualizar_data():
    """ Atualiza e retorna a data formatada para o dia atual. """
    return date.today().strftime("%d%m%Y")

def interagir_com_interface_superus(nome_do_arquivo):
    """ Realiza a interação com a interface do usuário para extrair dados. """
    # Abre os pedidos de compras
    pg.click(217,31, duration=0.5); sleep(2)
    pg.click(190,71, duration=0.5); sleep(2)
    
    while True:
        if pg.locateOnScreen(r"Imgs\notas_fiscais.png", confidence=0.9) is not None:
            break

    # Exporta para Excel
    pg.rightClick(276, 374, duration=0.5); sleep(5)
    pg.click(335, 880, duration=0.5)
    
    while True:
        pg.locateOnScreen(r"Imgs\meu computador.png"); sleep(3)
        break
    
    pg.click(365, 472, duration=0.5)
    
    while True:
        pg.locateOnScreen(r"Imgs\downloads.png"); sleep(3)
        break
    
    pg.doubleClick(502,356, duration=0.5, interval=0.5)
    
    while True:
        if pg.locateOnScreen(r"Imgs\salvar.png", confidence=0.9) is not None:
            break
        
    pg.press("tab", presses=2, interval=1)
    
    pg.write(nome_do_arquivo); pg.press("enter")

def listar_nfs_e_criar_df():
    """
    Lista os nomes dos arquivos PDF em diretórios específicos dentro de um diretório raiz e 
    cria um DataFrame com o nome do arquivo e a loja correspondente. 
    Exporta o DataFrame para um arquivo Excel.
    """
    # Definindo os diretórios
    dir_smj = r"G:\SCANNER_FISCAL SMJ"
    dir_stt = r"G:\SCANNER_FISCAL STT"
    dir_vix = r"G:\SCANNER_FISCAL"

    # Lista para armazenar informações dos arquivos
    arquivos_info = []

    # Função para adicionar arquivos de um diretório à lista
    def adicionar_arquivos(diretorio, loja):
        for arquivo in os.listdir(diretorio):
            if os.path.isfile(os.path.join(diretorio, arquivo)) and arquivo.lower().endswith(('.pdf', '.jpg')):
                # Verificando se 'boleto' não está no nome do arquivo
                if 'nf' in arquivo.lower():
                    # Separando o nome do arquivo de sua extensão
                    nome_arquivo, extensao = os.path.splitext(arquivo)
                    arquivos_info.append({"FORNECEDOR": nome_arquivo, "Loja": loja})

    # Adicionando arquivos de cada diretório
    adicionar_arquivos(dir_smj, "SMJ")
    adicionar_arquivos(dir_stt, "STT")
    adicionar_arquivos(dir_vix, "VIX")

    # Criando um DataFrame
    df = pd.DataFrame(arquivos_info)

    # Se a lista estiver vazia, cria um DataFrame vazio com os mesmos cabeçalhos
    if df.empty:
        df = pd.DataFrame(columns=["FORNECEDOR", "Loja"])

    # Exportando o DataFrame para Excel
    df.to_excel(r"Src\Notas_fiscais\Dashboard\nfs_recebidas.xlsx", index=False)

def aguardar_e_mover_download(arquivo_download, destino_arquivo, nome_arquivo):
    """ Aguarda a extração e move o arquivo para a pasta de destino. """
    while not os.path.exists(arquivo_download):
        sleep(5)

    if os.path.exists(arquivo_download):
        arquivo_antigo = os.path.join(destino_arquivo, nome_arquivo)
        
        if os.path.exists(arquivo_antigo):
            os.remove(arquivo_antigo); sleep(5)
            
        shutil.copy(arquivo_download, destino_arquivo)
    
    arquivo_2 = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Src\Notas_fiscais\Dashboard\recebimento_do_dia.xlsx"
    destino_2 = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Src\Notas_fiscais\Dashboard"

    if os.path.exists(arquivo_download):
        arquivo_antigo = os.path.join(destino_2, nome_arquivo)
        
        if os.path.exists(arquivo_antigo):
            os.remove(arquivo_antigo); sleep(5)
            
        shutil.move(arquivo_download, destino_2)

    pg.press("f9")  # Sai para o menu principal

    # Caminho para o arquivo .bat
    caminho_do_arquivo = r'F:\COMPRAS\Automações.Compras\Automações\Superus\Src\Notas_fiscais\Dashboard\daily_commit.bat'

    # Executar o arquivo .bat
    os.system(caminho_do_arquivo)

def extrair_notas_fiscais():    
    interagir_com_interface_superus("recebimento_do_dia") 
    listar_nfs_e_criar_df()
    aguardar_e_mover_download(r"C:\Users\automacao.compras\Downloads\recebimento_do_dia.xlsx", r"F:\BI\Bases", "recebimento_do_dia.xlsx")