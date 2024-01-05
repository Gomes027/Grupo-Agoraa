import os
import re
import csv
import sys
import shutil
import keyboard
import pyperclip
from time import sleep
import pyautogui as pg
import pygetwindow as gw
from PIL import ImageGrab

# CONSTANTES
DIR_RELATORIOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Relatórios"
DIR_GUSTAVO = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Imgs\Gustavo"
DIR_JAIDSON = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Imgs\Jaidson"

# FUNÇÕES AUXILIARES
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

def pegar_dados():
    estado_nome = True
    while True:
        dados_area_transferencia = pyperclip.paste()
        if dados_area_transferencia:
            if estado_nome:
                dados = dados_area_transferencia
                estado_nome = False
                return dados

# Função para aguardar até que a área de transferência tenha conteúdo
def aguardar_conteudo_area_transferencia(atalho, tempo_espera=30):
    pyperclip.copy("")  # Limpa a área de transferência
    conteudo_inicial = pyperclip.paste()  # Guarda o conteúdo inicial (vazio)

    if atalho == "Nome do Pedido":
        pg.hotkey("ctrl", "shift", "l")
    elif atalho == "Dados do Pedido":
        pg.hotkey("ctrl", "shift", "r")
    elif atalho == "CSV do Pedido":
        pg.hotkey("ctrl", "shift", "t") 
    elif atalho == "Imagem do Pedido":
        pg.hotkey("ctrl", "shift", "s") 

    tempo_decorrido = 0
    while tempo_decorrido < tempo_espera:
        sleep(1)  # Aguarda um segundo
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

        # Reescrever o arquivo com as linhas válidas
        with open(caminho_completo, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(linhas_validas)

    estado_nome_arquivo = True
    dados_area_transferencia = aguardar_conteudo_area_transferencia("Nome do Pedido")

    while True:
        if dados_area_transferencia:
            if estado_nome_arquivo:
                nome_arquivo = limpar_dados(dados_area_transferencia)
                estado_nome_arquivo = False
                print(f"Nome do arquivo CSV definido como: {nome_arquivo}")
                aguardar_conteudo_area_transferencia("CSV do Pedido")
                dados_area_transferencia = pyperclip.paste()
            else:
                salvar_em_csv(nome_arquivo, dados_area_transferencia)
                print(f"Dados salvos no arquivo {nome_arquivo}.csv")
                estado_nome_arquivo = True
                break

def salvar_imgs(caminho_imagens):
    # Capturar a imagem da área de transferência
    aguardar_conteudo_area_transferencia("Imagem do Pedido")

    try:
        imagem = ImageGrab.grabclipboard()
        if imagem:
            # Salvar a imagem no caminho especificado
            imagem.save(caminho_imagens, 'PNG')
            print(f"Imagem salva em: {caminho_imagens}")
        else:
            print("Não há imagem na área de transferência.")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    
    print("\n")

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

        # Realiza alguma ação com o nome
        print(f"Processando imagem: {nome_sem_extensao}")
    
    keyboard.wait("enter")

    for imagem in imagens:
        caminho_imagem = os.path.join(pasta, imagem)
        os.remove(caminho_imagem)

# FUNÇÕES PRINCIPAIS
def preencher_dados_excel(relatorio):
    # Ler e ordenar o relatório
    with open(relatorio, newline='') as arquivo:
        leitor = csv.reader(arquivo, delimiter=';')
        linhas_ordenadas = sorted(leitor, key=lambda row: row[2])  # Ordena pelo índice 3 (quarta coluna)

    loja_macro_anterior = None  # Variável para armazenar a loja_macro anterior

    for linha in linhas_ordenadas:
        fornecedor = linha[0]
        loja = linha[1]
        loja_macro = linha[2].strip()

        # Seleciona o Fornecedor
        pg.hotkey('ctrl', 'home')
        for _ in range(3):
            pg.hotkey('ctrl', 'up')
        pg.press('right')
        pg.press('down')
        pg.hotkey('alt', 'down')
        pg.press('space')
        pg.press('backspace')
        sleep(1)
        keyboard.write(fornecedor)
        sleep(0.5)

        pg.press('enter')
        sleep(5)

        # Usa o macro para selecionar as lojas
        if loja_macro != loja_macro_anterior:
            if loja_macro.strip() == 'M': # SMJ
                pg.hotkey("ctrl", "shift", "m")
            elif loja_macro.strip() == 'N': # STT
                pg.hotkey("ctrl", "shift", "n")
            elif loja_macro.strip() == 'B': # VIX
                pg.hotkey("ctrl", "shift", "b")
            elif loja_macro.strip() == 'C': # SMJ e STT
                pg.hotkey("ctrl", "shift", "c")
            elif loja_macro.strip() == 'X': # SMJ e VIX
                pg.hotkey("ctrl", "shift", "x")
            elif loja_macro.strip() == 'E': # SMJ, STT e MCP
                pg.hotkey("ctrl", "shift", "e")
            elif loja_macro.strip() == 'W': # SMJ, VIX e MCP
                pg.hotkey("ctrl", "shift", "w")
            elif loja_macro.strip() == 'V': # VIX, SMJ e STT
                pg.hotkey("ctrl", "shift", "v") 
            elif loja_macro.strip() == 'Z': # VIX, SMJ, STT e MCP
                pg.hotkey("ctrl", "shift", "z")
            else:
                print("Lojas erradas, a digitação sera encerrada!")
                sys.exit()

            sleep(15)
            
        loja_macro_anterior = loja_macro

        pg.hotkey('ctrl', 'shift', 'j')
        sleep(0.5)          

        pg.press('tab', presses=2)
        pg.press('enter')  
        sleep(0.5) 

        dados = aguardar_conteudo_area_transferencia("Dados do Pedido")
        elementos = dados.split()

        if len(elementos) >= 3:
            comprador = elementos[0]  # Primeiro elemento
            valor_total = elementos[1]  # Segundo elemento
            valor_minimo = elementos[2]  # Último elemento

            print("Fornecedor:", fornecedor)
            print("Loja:", loja)
            print("Comprador:", comprador)
            print("Valor Total:", valor_total)
            print("Valor Mínimo:", valor_minimo)
        else:
            print("Dados insuficientes")

        if valor_total == "#REF!":      
            pg.hotkey('ctrl', 'shift', 'k')
            nome_img = f"{fornecedor} {loja} - NÃO DEU PEDIDO"
                
        elif float(valor_total) < float(valor_minimo):            
            pg.hotkey('ctrl', 'shift', 'k')
            nome_img = f"{fornecedor} {loja} - NÃO DEU MÍNIMO R${valor_total}"

        else:              
            nome_img = f"{fornecedor} {loja} R${valor_total}"

        if valor_total != "#REF!":
            while True:
                processar_e_salvar_dados(r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\CSVs")
                break
        
        if comprador == "JAIDSON":
            caminho_completo = os.path.join(DIR_JAIDSON, f"{nome_img}.png")
        elif comprador == "GUSTAVO":
            caminho_completo = os.path.join(DIR_GUSTAVO, f"{nome_img}.png")

        salvar_imgs(caminho_completo)

# LOOP PRINCIPAL
if __name__ == "__main__":
    while True:
        arquivos_existentes = set()

        while True:
            arquivos = set(os.listdir(DIR_RELATORIOS))
            novos_arquivos = arquivos - arquivos_existentes
            novos_arquivos = sorted(novos_arquivos, key=lambda arquivo: os.path.getmtime(
                os.path.join(DIR_RELATORIOS, arquivo)))

            for novo_arquivo in novos_arquivos:
                if novo_arquivo.endswith(".csv"):
                    arquivo_completo = os.path.join(DIR_RELATORIOS, novo_arquivo)
                    relatorios_antigos = os.path.join(DIR_RELATORIOS, "Antigos")

                    abrir_janela("NOVO MODELO COMPRAS v10.20")
                    preencher_dados_excel(arquivo_completo)
                    processar_imagens(DIR_JAIDSON)
                    processar_imagens(DIR_GUSTAVO)
                    
                    # Remova o arquivo após processamento
                    try:
                        shutil.move(arquivo_completo, relatorios_antigos)
                    except:
                        pass
                
                else:
                    print(f"{novo_arquivo} não é um relatório de pedidos.")
                
            arquivos_existentes = arquivos
            sleep(1)