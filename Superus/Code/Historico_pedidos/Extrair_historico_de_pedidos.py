import os
import shutil
import subprocess
from time import sleep
import pyautogui as pg
from datetime import date

# Desabilita o Failsafe do Pyautogui
pg.FAILSAFE = False

def iniciar_superus():
    """
    Inicia o aplicativo Superus e aguarda até que esteja pronto para uso.
    """
    subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
    
    while True:
        if pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
            pg.click(532, 574, duration=0.5); sleep(1)
            pg.write("123456"); sleep(3)
            pg.press("enter", presses=2, interval=1)
            break
        
    while True:
        if pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
            break

def atualizar_data():
    """ Atualiza e retorna a data formatada para o dia atual. """
    return date.today().strftime("%d%m%Y")

def interagir_com_interface_superus(data_inicial, data_final):
    """ Realiza a interação com a interface do usuário para extrair dados. """
    # Abre os pedidos de compras
    pg.click(205,31, duration=0.5); sleep(1)
    pg.click(189,51, duration=0.5); sleep(5)
    
    # Insere a data inicial e final
    pg.press("tab", presses=7); pg.typewrite(data_inicial)
    pg.press("tab"); sleep(3)
    pg.typewrite(data_final); sleep(3)

    # Realiza a pesquisa
    pg.hotkey("alt", "f"); sleep(20)

    # Classifica os resultados
    pg.click(198, 214, duration=0.5, clicks=2, interval=2); sleep(10)
    pg.click(425, 212, duration=0.5); sleep(5)

    # Exporta para Excel
    pg.rightClick(); sleep(5)
    pg.press("up", presses=2); pg.press("enter")
    sleep(5); pg.write("historico de pedidos _oficial"); pg.press("enter")

def aguardar_e_mover_download(arquivo_download, pasta_destino_1, pasta_destino_2):
    """Aguarda a extração e move o arquivo para a pasta de destino, tentando até conseguir."""
    while not os.path.exists(arquivo_download):
        sleep(1)  # Aguarda o arquivo ser baixado
    
    # Define os nomes completos dos arquivos nas pastas de destino
    arquivo_antigo_1 = os.path.join(pasta_destino_1, "historico de pedidos _oficial.xlsx")
    arquivo_antigo_2 = os.path.join(pasta_destino_2, "historico de pedidos _oficial.xlsx")
    
    # Remove os arquivos antigos, se existirem
    if os.path.exists(arquivo_antigo_1):
        os.remove(arquivo_antigo_1)
        sleep(10)  # Aguarda um pouco após remover
    if os.path.exists(arquivo_antigo_2):
        os.remove(arquivo_antigo_2)
        sleep(10)  # Aguarda um pouco após remover
    
    sucesso = False
    while not sucesso:
        try:
            shutil.copy(arquivo_download, pasta_destino_2)  # Tenta copiar para o destino 2
            shutil.move(arquivo_download, pasta_destino_1)  # Tenta mover para o destino 1
            sucesso = True
            print("Extraído com Sucesso!\n")
        except PermissionError:
            print("Erro de permissão, tentando novamente em 5 segundos...")
            sleep(5)  # Espera 5 segundos antes de tentar novamente

def extrair_historico_pedidos():
    """ Função principal para extrair o histórico de pedidos. """
    print("Extraindo histórico de pedidos...")
    
    data_inicial = "01012022"  # Data fixa 
    data_final = atualizar_data()
    
    interagir_com_interface_superus(data_inicial, data_final)
    aguardar_e_mover_download(r"C:\Users\automacao.compras\Downloads\historico de pedidos _oficial.xlsx", r"F:\COMPRAS", r"F:\BI\Bases")
    pg.press("f9", presses=2, interval=5)

if __name__ == "__main__":
    iniciar_superus()
    extrair_historico_pedidos()