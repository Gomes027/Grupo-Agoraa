import os
import csv
import sys
import ctypes
import shutil
import locale
import chardet
import keyboard
import pytesseract
import pyautogui as pg
from io import BytesIO
from time import sleep
from PIL import Image, ImageEnhance
from datetime import date, timedelta

# Configuração do Tesseract
PATCH_TESSERACT = r'C:\Users\automacao.compras\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
pytesseract.pytesseract.tesseract_cmd = PATCH_TESSERACT

# Configuração do locale para usar o formato brasileiro de moeda
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

# FUNÇÔES CORRETIVAS   
def criar_caminho_imagem(nome_imagem):
    return os.path.join(r"C:\Users\automacao.compras\Pictures\Imgs\Pedidos de Compra", f"{nome_imagem}.png")

# Criando os caminhos das imagens
proxima_visita = criar_caminho_imagem("data_proxima_visita")
bloqueado = criar_caminho_imagem("bloqueado")
fora_do_mix = criar_caminho_imagem("fora_do_mix")
produto_ja_existe = criar_caminho_imagem("produto_ja_existe")
preço_de_venda = criar_caminho_imagem("venda")
custo_unitario = criar_caminho_imagem("unitario")
quantidade_obrigatoria = criar_caminho_imagem("obrigatorio")
recadastramento = criar_caminho_imagem("fornecedor_recadastramento")
preecher_comprador_img = criar_caminho_imagem("preencher_comprador")

def data_proxima_visita(data_da_proxima_visita):
    """ Corrige a data da proxima visita se for menor que a data atual. """
    pg.press("enter")
    pg.typewrite(data_da_proxima_visita)
    sleep(2)
    pg.hotkey("alt", "p")
    sleep(1)
    pg.click(35, 933, duration=1)
    sleep(2)
    sleep_por_imagem("janela_digitar_produtos.png")
    
def produto_bloqueado_ou_fora_do_mix():
    """ Se o produto estiver bloqueado, fora do mix, ou já existir no pedido, continua com o proximo produto. """
    pg.press("enter")
    sleep(1)
    pg.press("tab", presses=5)
    
def sem_preço_de_venda_ou_custo_unitário():
    """ Corrige a falta de preço de venda ou custo unitário. """
    pg.press("enter")
    sleep(2)
    pg.doubleClick(344, 573, duration=0.5)
    sleep(1)
    pg.typewrite("99")
    pg.press("enter")
    sleep(1)
    pg.hotkey("alt", "o")

# FUNÇÔES AUXILIARES
def atualizar_data():
    DATA_ATUAL = date.today()
    DATA_FORMATADA = DATA_ATUAL.strftime("%d/%m/%Y")
    DATA_PEDIDO = DATA_ATUAL.strftime("%d.%m")
    DATA_CORRECAO = (DATA_ATUAL + timedelta(days=7)).strftime("%d%m%Y")
    return DATA_FORMATADA, DATA_PEDIDO, DATA_CORRECAO

def extrair_fornecedor_loja_comprador(arquivo_completo):
    """ Função que Extrai os dados dos pedidos. """
     # Remove o caminho e deixa somente o nome do arquivo
    nome_arquivo, _ = os.path.splitext(os.path.basename(arquivo_completo))
    palavras = nome_arquivo.split() # Remove espaços inuteis

    if len(palavras) >= 3: # Lê o nome se o arquivo possuir mais que 3 elementos
        fornecedor = " ".join(palavras[:-2]) # Remove os dois ultimos elementos do nome e seleciona o restante
        loja = palavras[-2] # Seleciona o penúltimo elemento
        comprador = palavras[-1] # Seleciona o último elemento
        return fornecedor, loja, comprador
    else:
        return None, None, None
    
def sleep_por_imagem(imagem):
    """ Função que aguarda um elemento ou pagina carregar. """
    while True: # Aguarda até alguma pagina ou elemento carregar, via imagem
        imagem_completa = os.path.join(r"C:\Users\automacao.compras\Pictures\Imgs\Pedidos de Compra", imagem)
        localizacao = pg.locateOnScreen(imagem_completa, confidence=0.9)
        if localizacao is not None:
            break  
        
def remover_produtos_sem_quantidade(arquivo_completo):
    # Detectar a codificação do arquivo
    with open(arquivo_completo, 'rb') as file:
        encoding = chardet.detect(file.read())['encoding']

    # Ler as linhas do arquivo e armazenar linhas que não têm 0 no segundo item
    linhas_validas = []
    with open(arquivo_completo, mode='r', newline='', encoding=encoding) as file:
        reader = csv.reader(file, delimiter=';')
        for linha in reader:
            if linha[1] != '0':
                linhas_validas.append(linha)

    # Reescrever o arquivo com as linhas válidas
    with open(arquivo_completo, mode='w', newline='', encoding=encoding) as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerows(linhas_validas)
    
def numero_do_pedido():
    screenshot = pg.screenshot(region=(73, 63, 75, 25))
    imagem_bytesio = BytesIO()
    screenshot.save(imagem_bytesio, format='PNG')
    imagem = Image.open(imagem_bytesio)

    # Calcula o novo tamanho
    novo_tamanho = (int(imagem.width * 2), int(imagem.height * 2))

    # Redimensiona a imagem com o filtro LANCZOS
    imagem = imagem.resize(novo_tamanho, Image.LANCZOS)

    # Se necessário, aumente o contraste aqui
    enhancer = ImageEnhance.Contrast(imagem)
    imagem = enhancer.enhance(2)

    num_do_pedido = pytesseract.image_to_string(imagem, config='--psm 6', lang='por')
    num_do_pedido = num_do_pedido.strip().replace('\n', '').replace('\r', '')

    # Tentativa de conversão para inteiro
    try:
        num_do_pedido = int(num_do_pedido)

    except ValueError:
        print("Erro ao converter o número do pedido para inteiro.")
        num_do_pedido = 0
    
    return num_do_pedido

def valor_do_pedido():
    screenshot = pg.screenshot(region=(305, 126, 75, 25))
    imagem_bytesio = BytesIO()
    screenshot.save(imagem_bytesio, format='PNG')
    imagem = Image.open(imagem_bytesio)

    # Calcula o novo tamanho
    novo_tamanho = (int(imagem.width * 2), int(imagem.height * 2))

    # Redimensiona a imagem com o filtro LANCZOS
    imagem = imagem.resize(novo_tamanho, Image.LANCZOS)

    # Se necessário, aumente o contraste aqui
    enhancer = ImageEnhance.Contrast(imagem)
    imagem = enhancer.enhance(2)

    valor_do_pedido = pytesseract.image_to_string(imagem, config='--psm 6', lang='por')
    valor_do_pedido = valor_do_pedido.strip().replace('\n', '').replace('\r', '')

    # Tentativa de conversão para float
    try:
        valor_do_pedido = float(valor_do_pedido.replace(",", "."))
        valor_do_pedido_formatado = locale.currency(valor_do_pedido, grouping=True)
    except ValueError:
        print("Erro ao converter o número do pedido para float.")
        valor_do_pedido_formatado = "R$ 0,00"
    
    return valor_do_pedido_formatado

def desativar_capslock():
    return ctypes.windll.user32.GetKeyState(0x14) & 1 != 0

def verificar_tipo_arquivo(arquivo_completo):
    with open(arquivo_completo, 'rb') as file:
        encoding = chardet.detect(file.read())['encoding']

    with open(arquivo_completo, newline='', encoding=encoding) as csvfile:
        reader = csv.reader(csvfile, delimiter=';')  # Altere o delimitador se necessário
        for linha in reader:
            if linha:  # Verifica se a linha não está vazia
                if len(linha) == 3:
                    return "compras"
                elif len(linha) == 4:
                    return "cotação"
                else:
                    return "formato desconhecido"  # Caso não tenha nem 3 nem 4 colunas
    return "arquivo vazio"  # Caso todas as linhas estejam vazias

# Função para iniciar a captura dos logs do terminal
def iniciar_captura_logs(nome_arquivo_log):
    original_stdout = sys.stdout  # Salva a saída padrão original
    log_file = open(nome_arquivo_log, 'w')  # Abre o arquivo de log para escrita
    sys.stdout = log_file  # Redireciona stdout para o arquivo de log
    return original_stdout, log_file

# Função para finalizar a captura dos logs do terminal
def finalizar_captura_logs(original_stdout, log_file):
    sys.stdout = original_stdout  # Restaura a saída padrão original
    log_file.close()  # Fecha o arquivo de log
        
# FUNÇÔES PRINCIPAIS
def preencher_informações_pedido(fornecedor, loja, comprador, data_da_proxima_visita, arquivo_completo, tipo_pedido):
    """ Função Responvel por preencher as principais informações do pedido. """
    # Novo Pedido de compra
    pg.click(234, 707, duration=1); sleep(5)
    pg.press('F2')
    
    while True:
        if pg.locateOnScreen(r"C:\Users\automacao.compras\Pictures\Imgs\Pedidos de Compra\pedido_compra.png", confidence=0.9) is not None:
            break

    # Preenche o Fornecedor
    pg.click(143, 100, duration=1); sleep(2)
    keyboard.write(fornecedor)
    pg.press('enter')
    pg.hotkey('alt', 'o')
    sleep(1)
    
    if pg.locateOnScreen(recadastramento) is not None: 
        pg.press("tab")
        pg.press("enter")
    
    pg.press('tab')
    sleep(1)

    # Preenche o Comprador
    if comprador.lower() == "jaidson":
        pg.write(comprador)
        pg.press('enter')
    elif comprador.lower() == "gustavo":
        pg.write(comprador)
        pg.press('enter')
    else:
        print(f"Comprador inválido: {comprador}. A digitação será encerrada.")
        sys.exit()

    # Preenche a Loja
    pg.click(665, 67, duration=1)
    if loja.lower() == "smj":
        pg.press('up', presses=6)
        pg.press('enter')
    elif loja.lower() == "stt":
        pg.press('up', presses=6)
        pg.press('down')
        pg.press('enter')
    elif loja.lower() == "vix":
        pg.press('up', presses=6)
        pg.press('down', presses=2)
        pg.press('enter')
    elif loja.lower() == "mcp":
        pg.press('up', presses=6)
        pg.press('down', presses=3)
        pg.press('enter')
    else:
        print(f"Loja inválida: {loja}. A digitação será encerrada.")
        sys.exit()

    # Frete e prazo de pagamento
    pg.hotkey("alt", "f")
    pg.click(205, 240, duration=1)
    pg.hotkey('alt', 'i')
    pg.click(373, 187, duration=1)
    pg.typewrite('28')
    pg.press('enter')
    sleep(1)

    # guia "Produtos"
    pg.hotkey('alt', 'p')
    sleep(1)

    # Abre a Janela de Digitar Produtos
    pg.click(35, 933, duration=1)
    sleep(2)

    if pg.locateOnScreen(preecher_comprador_img) is not None:
        # Preenche o Comprador
        if comprador.lower() == "jaidson":
            pg.press('enter'); sleep(3)
            pg.write(comprador); sleep(1)
            pg.press('enter')
        elif comprador.lower() == "gustavo":
            pg.press('enter'); sleep(3)
            pg.write(comprador); sleep(1)
            pg.press('enter')
        else:
            print(f"Comprador inválido: {comprador}. A digitação será encerrada.")
            sys.exit()
        
        # Abre a Janela de Digitar Produtos
        pg.click(35, 933, duration=1)
        sleep(2)
        
    if pg.locateOnScreen(proxima_visita) is not None:
        # Corrige o Erro
        data_proxima_visita(data_da_proxima_visita)
    else:
        # Aguarda a janela de digitar produtos abrir
        sleep_por_imagem("janela_digitar_produtos.png")
        
        # Digita o controle de pedido
        pg.doubleClick(381, 206, duration=0.5)
        pg.write("81569")
        pg.press("enter")
        sleep(2)
        
        if pg.locateOnScreen(produto_ja_existe) is not None:
            produto_bloqueado_ou_fora_do_mix()
        else:
            pg.doubleClick(734, 389, duration=0.5)
            pg.write("1")
            pg.doubleClick(343, 574, duration=0.5)
            pg.write("0,01")
            pg.press("enter")
            pg.hotkey("alt", "o")
            sleep(2)

    if tipo_pedido == "compras":
        remover_produtos_sem_quantidade(arquivo_completo)

def digitar_produtos(tipo_pedido, arquivo_completo):
    with open(arquivo_completo) as arquivo:
        linhas_arquivo = arquivo.readlines()  # Lê os dados do arquivo


        for linha in linhas_arquivo:
            if tipo_pedido == "cotação":
                codigo, embalagem, quantidade, preço = linha.strip().split(";")
            elif tipo_pedido == "compras":
                codigo, quantidade, embalagem  = linha.strip().split(";")

            # Digita o codigo do produto
            pg.write(codigo); pg.press("enter"); sleep(0.5)
            
            if pg.locateOnScreen(bloqueado) is not None: 
                produto_bloqueado_ou_fora_do_mix() # Corrige o erro de produto bloqueado
                print(f"Produto {codigo} bloqueado.")

            elif pg.locateOnScreen(fora_do_mix) is not None:
                produto_bloqueado_ou_fora_do_mix() # Corrige o erro de produto fora do mix
                print(f"Produto {codigo} fora do mix.")

            elif pg.locateOnScreen(produto_ja_existe) is not None:
                produto_bloqueado_ou_fora_do_mix() # Corrige o erro de produto já existe nesse pedido
            
            else:
                pg.press("enter"); pg.write(quantidade); pg.press("enter") # Digita a Quantidade

                if tipo_pedido == "cotação":
                    pg.doubleClick(343, 574, duration=0.5)
                    pg.write(preço)
                    pg.press("enter")
                
                pg.hotkey("alt", "o"); sleep(0.5)

                if pg.locateOnScreen(preço_de_venda) is not None:
                    sem_preço_de_venda_ou_custo_unitário() # Corrige o erro de produto sem preço de venda
                    print(f"Produto {codigo} sem preço de venda.")

                elif pg.locateOnScreen(custo_unitario) is not None:
                    sem_preço_de_venda_ou_custo_unitário() # Corrige o erro de produto sem preço unitário
                    print(f"Produto {codigo} sem preço unitário.")
            
            # Espera a proxima iteração carregar
            sleep(0.5)

        # Fecha a janela de digitar produtos
        pg.press("esc")
        sleep(3)
        
def salvar_pedido(fornecedor, loja, num_do_pedido, data_pedido, tipo_pedido, valor_formatado, data_historico):
    pg.press("f11")
    sleep_por_imagem("salvar_pedido.png")
    
    pg.hotkey("alt", "o")

    sleep_por_imagem("salvar_pdf.PNG")
    sleep(2)
    
    pg.click(465, 39, duration=0.5)
    sleep(2)

    if desativar_capslock():
        pg.press('capslock')
    
    dados = f"{fornecedor} {loja} {num_do_pedido} {data_pedido}"
    print(dados)
    print("\n")
    
    arquivo_com_extensão = (f"{dados}.pdf")
    arquivo_downloads = os.path.join(r"C:\Users\automacao.compras\Downloads", arquivo_com_extensão)
    
    pg.typewrite(dados)
    sleep(1)
    
    pg.press("enter")
    sleep(1)
    
    pg.press("f9")
    sleep(2)

    pg.press("esc")
    sleep(1)

    pg.press("f9")
    sleep(2)
    
    pg.press("f9")
    sleep(2)

    if tipo_pedido == "compras":
        dados_lista = []
        
        dados_lista.append([num_do_pedido, fornecedor, loja, valor_formatado, data_historico])

        with open(r"Outros\Historico_compras.csv", 'a', newline='', encoding='utf-8') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            arquivo_vazio = arquivo.tell() == 0

            if arquivo_vazio:
                escritor.writerow(['Numero do Pedido', 'Fornecedor', 'Loja', 'Valor do Pedido', 'Data'])

            for linha in dados_lista:
                escritor.writerow(linha)
    
    return arquivo_downloads

def digitar_pedido(arquivo_completo, tipo_pedido, novo_arquivo):
    """ Função responsavel por realizar os pedidos. """
    DATA_FORMATADA, DATA_PEDIDO, DATA_CORRECAO = atualizar_data()

    fornecedor, loja, comprador = extrair_fornecedor_loja_comprador(arquivo_completo)
    if fornecedor and loja and comprador:
        print(f"Arquivo: {novo_arquivo}")
        print(f"Fornecedor: {fornecedor}, Loja: {loja}, Comprador: {comprador}")
        sleep(1)
        
        preencher_informações_pedido(fornecedor, loja, comprador, DATA_CORRECAO, arquivo_completo, tipo_pedido)
    else:
        print(f"Informações do pedido {novo_arquivo} não encontradas.")
        sys.exit()

    digitar_produtos(tipo_pedido, arquivo_completo)
    
    num_do_pedido = numero_do_pedido()

    valor_formatado = valor_do_pedido()

    arquivo_downloads = salvar_pedido(fornecedor, loja, num_do_pedido, DATA_PEDIDO, tipo_pedido, valor_formatado, DATA_FORMATADA)
    
    if tipo_pedido == "compras":
        shutil.move(arquivo_downloads, r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs")
    if tipo_pedido == "cotação":
        shutil.move(arquivo_downloads, r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs")

# Abre o Superus
os.system("cls")
pg.hotkey('win', '3')
sleep(1)