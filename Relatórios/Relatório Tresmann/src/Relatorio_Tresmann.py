import os
import time
import shutil
import ctypes
import zipfile
import asyncio
from time import sleep
import pyautogui as pg
from pyppeteer import launch
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
        pg.press('capslock')

def esperar_planilha_carregar(imagem1, imagem2):
    """ Aguarda as Funcionalidades da planilha carregar. """
    while True:
        localizacao1 = pg.locateOnScreen(imagem1, confidence=0.9)
        localizacao2 = pg.locateOnScreen(imagem2, confidence=0.9)

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
async def baixar_arquivo(email, senha, DIR_DOWNLOAD):
    browser = await launch(headless=False, args=['--start-maximized'])
    page = await browser.newPage()

    # Configura o diretório de download
    await page._client.send('Page.setDownloadBehavior', {'behavior': 'allow', 'downloadPath': DIR_DOWNLOAD})

    max_attempts = 3
    attempt = 0

    while attempt < max_attempts:
        try:
            await page.goto('http://webmail.agoraa.com.br/', {'timeout': 60000})
            break  # Se a página carregar, sair do loop
        except pyppeteer.errors.TimeoutError:
            print(f"Timeout ao tentar acessar a página. Tentativa {attempt + 1} de {max_attempts}.")
            attempt += 1
            if attempt < max_attempts:
                await asyncio.sleep(5)  # Espera 5 segundos antes de tentar novamente

    if attempt == max_attempts:
        print("Não foi possível acessar a página após várias tentativas.")

    # Preenche o formulário de login
    await page.type('#user', email)
    await page.type('#pass', senha)
    await page.click('#login_submit')

    max_attempts = 3
    attempt = 0
    navegacao_bem_sucedida = False

    while attempt < max_attempts and not navegacao_bem_sucedida:
        try:
            await page.waitForNavigation({'timeout': 60000})
            navegacao_bem_sucedida = True
        except pyppeteer.errors.TimeoutError:
            print(f"Timeout ao esperar pela navegação. Tentativa {attempt + 1} de {max_attempts}.")
            attempt += 1
            if attempt < max_attempts:
                await asyncio.sleep(5)  # Espera 5 segundos antes de tentar novamente

    if not navegacao_bem_sucedida:
        print("Falha ao navegar para a próxima página após o login.")
        return await browser.close()
    
    await page.type('#quicksearchbox', 'relatório diário')
    await page.keyboard.press('Enter')

    # Espera o elemento ser carregado na página
    await page.waitForSelector('table#messagelist tbody tr:nth-child(1)')
    
    # Primeiro, encontre o elemento
    element_handle = await page.querySelector('table#messagelist tbody tr:nth-child(1)')

    # Em seguida, realize um clique duplo
    if element_handle:
        await element_handle.click()  # Primeiro clique
        await element_handle.click()  # Segundo clique

    # Antes de clicar, espera pelo elemento usando XPath
    await page.waitForXPath('//*[@id="attach2"]/a[1]/span[1]')

    # Obter o elemento e clicar nele
    element_handle = await page.xpath('//*[@id="attach2"]/a[1]/span[1]')
    if element_handle:
        await element_handle[0].click()

    # Espera o download ser concluído
    timeout = 480
    start_time = time.time()
    await asyncio.sleep(5)  # Espera inicial para dar tempo ao download de começar

    while True:
        still_downloading = any(".crdownload" in filename for filename in os.listdir(DIR_DOWNLOAD))
        if not still_downloading:
            print("Download concluído.")
            break

        elapsed_time = time.time() - start_time
        if elapsed_time > timeout:
            print("Tempo de espera excedido para o download.")
            break

        print(f"Aguardando o download... {elapsed_time:.2f} segundos passados")
        await asyncio.sleep(1)

    await browser.close()

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

    while True:
        if pg.locateOnScreen(r"Imgs\planilha_aberta2.png", confidence=0.7):
            sleep(5)
            break

        sleep(1)

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
    sleep(20)

    # Verificar o estado do Caps Lock antes de salvar
    verificar_caps_lock()

    # Salvar a planilha
    pg.write('relatorio_tresmann')
    sleep(10)

    pg.press('tab'); sleep(1)
    pg.press('down'); sleep(1)
    pg.press('up', presses=30); sleep(1)
    pg.press('enter'); sleep(5)

    pg.hotkey("alt", "s")
    # Aguardar um tempo para garantir o salvamento
    sleep(60)

    pg.hotkey('alt', 'f4')
    sleep(20)

def main():
    try:
        limpar_diretorio(DIR_DOWNLOAD)
        print("Baixando Arquivo...")
        # Executa a função baixar_arquivo
        asyncio.get_event_loop().run_until_complete(baixar_arquivo(email, senha, DIR_DOWNLOAD))

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

if __name__ == "__main__":
    main()