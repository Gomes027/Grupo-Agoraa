import os
import sys
import shutil
import subprocess
import pyautogui as pg
from time import sleep

# Desabilita o Failsafe do Pyautogui
pg.FAILSAFE = False

def iniciar_superus():
    """
    Inicia o aplicativo Superus e aguarda até que esteja pronto para uso.
    """
    subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
    
    while True:
        if pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
            pg.write("123456"); sleep(3)
            pg.press("enter", presses=2, interval=1)
            break
        
    while True:
        if pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
            break

def interagir_com_interface_superus(select, data_inicial, data_final, nome_do_arquivo):
    # Abre a execução de selects
    pg.click(270,31, duration=0.5); sleep(1)
    pg.click(257,46, duration=0.5); sleep(1)

    # Aguarda a aba consultas carregar
    while True:
        if pg.locateOnScreen(r"imgs\executar_consultas.png") is not None:
            break
        
    pg.click(415, 247, duration=0.5)

    # Insere e executa o select
    sleep(5)
    pg.typewrite(select); sleep(1)
    pg.press("enter"); sleep(5)
    pg.press("f8"); sleep(3)

    if data_inicial is not None:
        pg.write(data_inicial); sleep(1)
        pg.press("down"); sleep(1)
        pg.write(data_final); sleep(1)
        pg.press("enter", presses=2, interval=1); sleep(1)

    # Aguardar select
    while True:
        if pg.locateOnScreen(r"imgs\executando_select.png") is None:
            sleep(10)
            break

    # Exportar para excel
    pg.rightClick(808, 567, duration=0.5); sleep(3)
    pg.click(863, 810, duration=0.5)
    pg.locateOnScreen(r"Imgs\explorer.png"); sleep(3)
    
    nome_arquivo_sem_extensao = os.path.splitext(nome_do_arquivo)[0]
    pg.typewrite(nome_arquivo_sem_extensao); sleep(3); pg.press("enter")
    
def aguardar_e_mover_download(arquivo_download, pasta_destino, nome_do_arquivo):
    while True:
        try:
            # Aguarda até que o arquivo esteja disponível
            while not os.path.exists(arquivo_download):
                sleep(5)
            
            # Remove o arquivo antigo, se existir
            arquivo_antigo = os.path.join(pasta_destino, nome_do_arquivo)
            if os.path.exists(arquivo_antigo):
                os.remove(arquivo_antigo)
            
            # Tenta mover o arquivo
            shutil.move(arquivo_download, pasta_destino); sleep(3)
            pg.press("f9")  # Sai para o menu principal

            print("Arquivo salvo!")
            print("\n")

            # Sai do loop se o arquivo foi movido com sucesso
            break

        except Exception as e:
            print(f"Erro ao mover o arquivo: {e}")
            sleep(5)

# Função para Extração dos dados
def extrair_dados(tipo_de_select, data_inicio, data_fim, nome_do_arquivo):
    dir_arquivo = os.path.join(r"C:\Users\automacao.compras\Downloads", nome_do_arquivo)

    print(f"Extraindo dados para o select {tipo_de_select}...")
    interagir_com_interface_superus(tipo_de_select, data_inicio, data_fim, nome_do_arquivo)
    aguardar_e_mover_download(dir_arquivo, r"F:\BI\Bases", nome_do_arquivo)

# Função Principal, que executa todos os selects
def executar_selects():
    # Dados para extração
    dados_para_extracao = [
        ("1350", "01012022", "31122024", "historico_de_recebimento.xlsx"),
        ("1610", "01012022", "31122024", "historico_inventario.xlsx"),
        ("1590", None, None, "relatorio_vendas.xlsx"),
        ("1570", None, None, "relatorio_precos_4_lojas.xlsx")
    ]

    # Executar Extração para cada conjunto de dados
    for select, data_inicio, data_fim, nome_arquivo in dados_para_extracao:
        extrair_dados(select, data_inicio, data_fim, nome_arquivo)
        sleep(5)
        
if __name__ == "__main__":
    iniciar_superus()
    executar_selects()