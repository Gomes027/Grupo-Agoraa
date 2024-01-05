import os
import time
import shutil
import ctypes
import zipfile
import schedule
from time import sleep
import pyautogui as pg
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from config import email, senha

# Desabilita o Failsafe do Pyautogui
pg.FAILSAFE = False

# Funções Auxiliares
def limpar_diretorio(dir_path):
    """ Remove todos os arquivos em um diretório especificado. """
    for filename in os.listdir(dir_path):
        file_path = os.path.join(dir_path, filename)
        if os.path.isfile(file_path) or os.path.islink(file_path):
            os.unlink(file_path)

def verificar_caps_lock():
    """ Verifica se o Caps Lock está ativado e desativa se necessário. """
    if ctypes.windll.user32.GetKeyState(0x14) & 1 != 0:
        print("O Caps Lock está ativado. Desativando...")
        pg.press('capslock')
        print("Caps Lock desativado.")
    else:
        print("O Caps Lock não está ativado.")

def esperar_planilha_carregar(imagem1, imagem2):
    """ Aguarda as Funcionalidades da planilha carregar. """
    while True:
        localizacao1 = pg.locateOnScreen(imagem1)
        localizacao2 = pg.locateOnScreen(imagem2)

        if localizacao1 is not None:
            break
        elif localizacao2 is not None:
            break
            
        time.sleep(1)

def mover_arquivo(origem, destino_1, destino_2, remover_1, remover_2, extensao):
    """ Move um arquivo com a extensão especificada do diretório de origem para o destino. """
    for arquivo in os.listdir(origem):
        if arquivo.endswith(extensao):
            caminho_arquivo = os.path.join(origem, arquivo)
            if os.path.exists(remover_1):
                os.remove(remover_1)
                
            if os.path.exists(remover_2):
                os.remove(remover_2)
                
            shutil.copy(caminho_arquivo, destino_2)
            shutil.move(caminho_arquivo, destino_1)

# Funções Principais
def baixar_arquivo():
    """ Acessa o email e baixa o relatório tresmann mais recente. """
    limpar_diretorio(DIR_DOWNLOAD)
    pg.moveTo(500,500)

    chrome_options = Options()
    prefs = {"download.default_directory" : DIR_DOWNLOAD}
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("start-maximized")

    driver = webdriver.Chrome(options=chrome_options)
    driver.get("http://webmail.agoraa.com.br/")
    driver.implicitly_wait(20)

    driver.find_element('xpath', '//*[@id="user"]').send_keys(email)
    driver.find_element('xpath', '//*[@id="pass"]').send_keys(senha)
    driver.find_element('xpath', '//*[@id="login_submit"]').click()
    driver.implicitly_wait(20)

    driver.find_element('xpath', '//*[@id="quicksearchbox"]').send_keys("relatório diário" + Keys.ENTER )
    driver.implicitly_wait(20)

    elemento_para_duplo_clique = driver.find_element('xpath', '//table[@id="messagelist"]/tbody/tr[1]')
    acoes = ActionChains(driver)
    acoes.double_click(elemento_para_duplo_clique).perform()
    driver.implicitly_wait(20)

    driver.find_element('xpath', '//*[@id="attach2"]/a[1]/span[1]').click()
    sleep(5)

    timeout = 480
    tempo_inicial = time.time()

    while True:
        ainda_baixando = any(".crdownload" in filename for filename in os.listdir(DIR_DOWNLOAD))
        if not ainda_baixando:
            break

        if time.time() - tempo_inicial > timeout:
            print("Tempo de espera excedido para o download.")
            break

        time.sleep(1)

    driver.quit()

def extrair_e_abrir_arquivo():
    """ Realiza a extração e abertura do arquivo. """
    for arquivo in os.listdir(DIR_DOWNLOAD):
        if arquivo.endswith('.zip'):
            caminho_zip = os.path.join(DIR_DOWNLOAD, arquivo)
            with zipfile.ZipFile(caminho_zip, 'r') as zip_ref:
                zip_ref.extractall(DIR_DOWNLOAD)

    sleep(3)

    for arquivo in os.listdir(DIR_DOWNLOAD):
        if arquivo.endswith('.xls') or arquivo.endswith('.xlsx'):
            caminho_xls = os.path.join(DIR_DOWNLOAD, arquivo)
            os.system(f'start excel "{caminho_xls}"')

def ajustar_planilha():
    """ Realiza as manipulações necessárias na planilha. """
    esperar_planilha_carregar(r"Imgs\abrir_planilha1.png", r"Imgs\abrir_planilha2.png")
    pg.hotkey('alt', 's')

    esperar_planilha_carregar(r"Imgs\planilha_aberta.png", r"Imgs\planilha_aberta2.png")
    sleep(5)

    pg.hotkey('ctrl', 'home')
    sleep(2)
    pg.hotkey('ctrl', 'shift', 'l')
    sleep(1)
    pg.hotkey('alt', 'down')
    sleep(1)

    pg.press('down', presses=7)
    sleep(1)
    pg.typewrite('codigo')
    sleep(10)
    pg.press('enter')
    sleep(5)
    
    for _ in range(15):
        pg.press('down')
        sleep(1)
        pg.hotkey('ctrl', '-')
        sleep(1)
        pg.press('enter')
        sleep(1)
        pg.hotkey('ctrl', 'home')
    
    pg.hotkey('ctrl', 'shift', 'l')
    sleep(5)

    pg.press("f12")
    sleep(5)

    # Verificar o estado do Caps Lock antes de salvar
    verificar_caps_lock()


    # Salvar a planilha
    pg.write('relatorio_tresmann')
    sleep(5)

    pg.press('tab')
    pg.press('down')
    pg.press('up', presses=30)
    pg.press('enter')
    sleep(2)

    pg.hotkey("alt", "s")
    # Aguardar um tempo para garantir o salvamento
    sleep(60)

    pg.hotkey('alt', 'f4')
    sleep(20)

def main():
    try:
        print("Baixando Arquivo...")
        baixar_arquivo()

        print("Extraindo Arquivo...")
        extrair_e_abrir_arquivo()

        print("Ajustando Arquivo...")
        ajustar_planilha()

        print("Salvando Arquivo...")
        mover_arquivo(DIR_DOWNLOAD, r"F:\COMPRAS", r"F:\BI\Bases", r"F:\COMPRAS\relatorio_tresmann.xlsx", r"F:\BI\Bases\relatorio_tresmann.xlsx", '.xlsx')

        print("Arquivo Salvo.")
        print("\n")

    except Exception as e:
        with open("erro_log.txt", "a") as file:
            file.write(str(e))
            file.write("\n\n"),
        
        print("Erro ao Atualizar Relatório Tresmann...")
        
# Variável GLobal
DIR_DOWNLOAD = r"F:\COMPRAS\Automações.Compras\Automações\Relatórios\Relatório Tresmann\Planilhas"

# Agendar para executar todos os dias às 23:00 AM
schedule.every().day.at("22:00").do(main)

while True:
    schedule.run_pending()
    sleep(1)