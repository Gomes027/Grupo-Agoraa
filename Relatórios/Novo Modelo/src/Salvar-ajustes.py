import os
import re
import csv
import time
import keyboard
import pyperclip
import clipboard
import pyautogui as pg
from time import sleep

# Função para aguardar até que a área de transferência tenha conteúdo
def aguardar_conteudo_area_transferencia(atalho, tempo_espera=30):
    pyperclip.copy("")  # Limpa a área de transferência
    conteudo_inicial = pyperclip.paste()  # Guarda o conteúdo inicial (vazio)

    if atalho == "Nome do Pedido":
        pg.hotkey("ctrl", "shift", "l")
    elif atalho == "CSV do Pedido":
        pg.hotkey("ctrl", "shift", "t") 
    tempo_decorrido = 0
    while tempo_decorrido < tempo_espera:
        sleep(2)  # Aguarda um segundo
        tempo_decorrido += 1
        conteudo_atual = pyperclip.paste()
        if conteudo_atual != conteudo_inicial:
            return conteudo_atual # Retorna o novo conteúdo

    print("Tempo de espera excedido. Nenhum conteúdo novo na área de transferência.")
    return None

def processar_e_salvar_dados(caminho_da_pasta):
    """ Função que processa e salva os dados dos pedidos de compra, em formato CSV. """
    def limpar_dados(dados):
        nome_arquivo = re.sub(r'[<>:"/\\|?*]', '', dados).strip()
        return nome_arquivo

    def salvar_em_csv(nome_arquivo, dados):
        numeros = dados.split()
        linhas = [numeros[i:i + 3] for i in range(0, len(numeros), 3)]
        caminho_completo = os.path.join(caminho_da_pasta, nome_arquivo + ".csv")
        os.makedirs(os.path.dirname(caminho_completo), exist_ok=True)

        # Escrever as linhas do arquivo
        with open(caminho_completo, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=';')
            for linha in linhas:
                csv_writer.writerow(linha)
        
        # Ler as linhas do arquivo e armazenar linhas que não têm 0 no segundo item
        linhas_validas = []
        with open(caminho_completo, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file, delimiter=';')
            for linha in reader:
                if linha[1] != '0':
                    linhas_validas.append(linha)
                else:
                    pass

        # Reescrever o arquivo com as linhas válidas
        with open(caminho_completo, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(linhas_validas)

    keyboard.wait("ctrl+shift+q")

    estado_nome_arquivo = True
    dados_area_transferencia = aguardar_conteudo_area_transferencia("Nome do Pedido")

    while True:
        if dados_area_transferencia:
            if estado_nome_arquivo:
                nome_arquivo = limpar_dados(dados_area_transferencia)
                estado_nome_arquivo = False
                print(f"Arquivo: {nome_arquivo}")
                aguardar_conteudo_area_transferencia("CSV do Pedido")
                dados_area_transferencia = pyperclip.paste()
            else:
                salvar_em_csv(nome_arquivo, dados_area_transferencia)
                print(f"Salvo com sucesso!")
                print("\n")
                estado_nome_arquivo = True
                break

    pg.hotkey("ctrl", "shift", "s")

if __name__ == "__main__":
    while True:
        processar_e_salvar_dados(r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\CSVs")