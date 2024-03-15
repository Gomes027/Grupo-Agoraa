import os
import csv
import sys
import time
import ctypes
import shutil
import locale
import chardet
import logging
import datetime
import keyboard
import subprocess
import pytesseract
import pyautogui as pg
from io import BytesIO
from time import sleep
from PIL import Image, ImageEnhance
from datetime import date, timedelta

class ConfiguracoesIniciais:
    def __init__(self):
        self.patch_tesseract = r'C:\Users\automacao.compras\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
        self.caminho_imagens = r"C:\Users\automacao.compras\Pictures\Imgs\Pedidos de Compra"
        pytesseract.pytesseract.tesseract_cmd = self.patch_tesseract
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

class UtilitariosDeImagem:
    @staticmethod
    def criar_caminho_imagem(nome_imagem, caminho_base):
        return os.path.join(caminho_base, f"{nome_imagem}.png")

    @staticmethod
    def capturar_numero_do_pedido(region):
        imagem = UtilitariosDeImagem.capturar_e_processar_imagem(region)
        return pytesseract.image_to_string(imagem, config='--psm 6', lang='por').strip().replace('\n', '').replace('\r', '')

    @staticmethod
    def capturar_e_processar_imagem(region):
        screenshot = pg.screenshot(region=region)
        imagem_bytesio = BytesIO()
        screenshot.save(imagem_bytesio, format='PNG')
        imagem = Image.open(imagem_bytesio)
        imagem = imagem.resize((imagem.width * 2, imagem.height * 2), Image.LANCZOS)
        imagem = ImageEnhance.Contrast(imagem).enhance(2)
        return imagem
    
    @staticmethod
    def capturar_valor_do_pedido(region):
        """
        Captura e processa uma região específica da tela para extrair o valor do pedido.

        :param region: Uma tupla (x, y, largura, altura) que define a região da tela a ser capturada.
        :return: O valor do pedido como string.
        """
        imagem = UtilitariosDeImagem.capturar_e_processar_imagem(region)
        valor_do_pedido = pytesseract.image_to_string(imagem, config='--psm 6', lang='por').strip().replace('\n', '').replace('\r', '')
        return valor_do_pedido

    @staticmethod
    def sleep_por_imagem(imagem_caminho, confidence=0.9, timeout=30):
        """Aguarda até que uma imagem específica seja encontrada na tela ou até que o tempo limite seja atingido."""
        start_time = time.time()
        while True:
            current_time = time.time()
            if current_time - start_time > timeout:
                print("Timeout: Imagem não encontrada.")
                break
            localizacao = pg.locateOnScreen(os.path.join(r"C:\Users\automacao.compras\Pictures\Imgs\Pedidos de Compra", imagem_caminho), confidence=confidence)
            if localizacao is not None:
                print("Imagem encontrada.")
                break
            sleep(1)

class GerenciamentoDeArquivos:
    @staticmethod
    def extrair_fornecedor_loja_comprador(arquivo_completo):
        nome_arquivo, _ = os.path.splitext(os.path.basename(arquivo_completo))
        palavras = nome_arquivo.split()
        if len(palavras) >= 3:
            fornecedor = " ".join(palavras[:-2])
            loja, comprador = palavras[-2], palavras[-1]
            return fornecedor, loja, comprador
        else:
            logging.warning(f'Não foi possível extrair fornecedor, loja e comprador do arquivo: {nome_arquivo}')
            return None, None, None

    @staticmethod
    def remover_produtos_sem_quantidade(arquivo_completo):
        with open(arquivo_completo, 'rb') as file:
            encoding = chardet.detect(file.read())['encoding']
        linhas_validas = []
        with open(arquivo_completo, mode='r', newline='', encoding=encoding) as file:
            for linha in csv.reader(file, delimiter=';'):
                if linha[1] != '0':
                    linhas_validas.append(linha)
        with open(arquivo_completo, mode='w', newline='', encoding=encoding) as file:
            csv.writer(file, delimiter=';').writerows(linhas_validas)

class GerenciamentoSuperus:
    def iniciar_superus(self):
        subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
        self.aguardar_abertura()

    def aguardar_abertura(self):
        while not pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
            sleep(1)
        pg.click(572, 574, duration=0.5); sleep(1)
        pg.write("123456"); sleep(3)
        pg.press("enter", presses=2, interval=1)
        while not pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
            sleep(1)

    @staticmethod
    def fechar_processo():
        try:
            subprocess.run(['taskkill', '/f', '/im', "SUPERUS.EXE"], check=True)
        except subprocess.CalledProcessError:
            pass

# Classe principal de automação
class OperacoesDePedido:
    def __init__(self):
        self.configs = ConfiguracoesIniciais()
        self.util_imagem = UtilitariosDeImagem()
        self.ger_arquivos = GerenciamentoDeArquivos()
        
        # Passando as instâncias necessárias como argumentos para os construtores
        self.preenchimento_info = PreenchimentoPedido(self.util_imagem, self.configs)
        self.digitar_ped = DigitacaoProdutos(self.util_imagem, self.configs)
        self.salvar_ped = SalvamentoPedido(self.util_imagem, self.configs)

    def executar_automacao(self, arquivo_completo, tipo_pedido, novo_arquivo):
        logging.info(f'Processando pedido de {tipo_pedido}: {novo_arquivo}')
        
        DATA_FORMATADA, DATA_PEDIDO, DATA_CORRECAO = self.atualizar_data()
        fornecedor, loja, comprador = self.ger_arquivos.extrair_fornecedor_loja_comprador(arquivo_completo)

        if not fornecedor or not loja or not comprador:
            print(f"Informações do pedido {novo_arquivo} não encontradas.")
            return

        # Preencher as informações do pedido
        self.preenchimento_info.preencher_informacoes_pedido(fornecedor, loja, comprador, DATA_CORRECAO, arquivo_completo, tipo_pedido)

        if tipo_pedido == "compras":
            self.ger_arquivos.remover_produtos_sem_quantidade(arquivo_completo)

        self._inserir_controle_de_pedidos()

        # Digitar os produtos
        self.digitar_ped.digitar_produtos(tipo_pedido, arquivo_completo)

        # Obter número do pedido e valor formatado
        num_do_pedido = self.obter_numero_do_pedido()
        valor_formatado = self.obter_valor_do_pedido()

        # Salvar o pedido e mover o arquivo PDF conforme o tipo de pedido
        arquivo_downloads = self.salvar_ped.salvar_pedido(fornecedor, loja, num_do_pedido, DATA_PEDIDO, tipo_pedido, valor_formatado, DATA_FORMATADA)
        self.mover_arquivo_pedido(arquivo_downloads, tipo_pedido)

    def _inserir_controle_de_pedidos(self):
        # Aguarda a janela de digitar produtos abrir
        self.util_imagem.sleep_por_imagem("janela_digitar_produtos.png")
        
        # Digita o controle de pedido
        pg.doubleClick(381, 206, duration=0.5)
        pg.write("81569")
        pg.press("enter")
        sleep(2)
        
        if pg.locateOnScreen(self.util_imagem.criar_caminho_imagem("produto_ja_existe", self.configs.caminho_imagens)) is not None:
            self.digitar_ped.produto_bloqueado_ou_fora_do_mix()
        else:
            pg.doubleClick(734, 389, duration=0.5)
            pg.write("1")
            pg.doubleClick(343, 574, duration=0.5)
            pg.write("0,01")
            pg.press("enter")
            pg.hotkey("alt", "o")
            sleep(2)

    def atualizar_data(self):
        data_atual = date.today()
        data_formatada = data_atual.strftime("%d/%m/%Y")  # Data no formato dia/mês/ano
        data_pedido = data_atual.strftime("%d.%m")  # Data no formato dia.mês para uso em identificações específicas
        data_correcao = (data_atual + timedelta(days=7)).strftime("%d%m%Y")  # Uma semana a mais para a data de correção
        
        return data_formatada, data_pedido, data_correcao

    def obter_numero_do_pedido(self):
        # Exemplo: Definir a região da tela onde o número do pedido é exibido
        region = (73, 63, 75, 25)  # Exemplo de região; ajuste conforme necessário
        num_do_pedido = self.util_imagem.capturar_numero_do_pedido(region)
        try:
            num_do_pedido = int(num_do_pedido.strip().replace('\n', '').replace('\r', ''))
        except ValueError:
            print("Erro ao converter o número do pedido para inteiro.")
            num_do_pedido = 0
        
        return num_do_pedido

    def obter_valor_do_pedido(self):
        region = (305, 126, 75, 25)  # Exemplo de região; ajuste conforme necessário
        valor_do_pedido = self.util_imagem.capturar_valor_do_pedido(region)
        try:
            valor_do_pedido = float(valor_do_pedido.strip().replace(',', '.'))
            valor_formatado = f"R$ {valor_do_pedido:,.2f}".replace(',', ' temporário ').replace('.', ',').replace(' temporário ', '.')
        except ValueError:
            print("Erro ao converter o valor do pedido.")
            valor_formatado = "R$ 0,00"
        
        return valor_formatado

    def mover_arquivo_pedido(self, arquivo_downloads, tipo_pedido):
        destino_compras = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"
        destino_cotacoes = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"
        
        if tipo_pedido == "compras":
            destino = destino_compras
        elif tipo_pedido == "cotação":
            destino = destino_cotacoes
        else:
            print(f"Tipo de pedido desconhecido: {tipo_pedido}")
            return
        
        shutil.move(arquivo_downloads, destino)

class PreenchimentoPedido:
    def __init__(self, util_imagem, configs):
        self.util_imagem = util_imagem
        self.configs = configs

    def preencher_informacoes_pedido(self, fornecedor, loja, comprador, data_da_proxima_visita, arquivo_completo, tipo_pedido):
        """Função responsável por preencher as principais informações do pedido."""
        pg.click(234, 707, duration=1)
        sleep(5)
        pg.press('F2')
        
        while True:
            if pg.locateOnScreen(self.util_imagem.criar_caminho_imagem("pedido_compra", self.configs.caminho_imagens), confidence=0.9) is not None:
                break

        # Preenche o Fornecedor
        pg.click(143, 100, duration=1)
        sleep(2)
        keyboard.write(fornecedor)
        pg.press('enter')
        pg.hotkey('alt', 'o')
        sleep(3)
        
        recadastramento = self.util_imagem.criar_caminho_imagem("fornecedor_recadastramento", self.configs.caminho_imagens)
        if pg.locateOnScreen(recadastramento, confidence=0.9) is not None:
            pg.press("tab")
            pg.press("enter")
        
        pg.click(669,87, duration=3)
        sleep(1)

        # Preenche o Comprador
        self._preencher_comprador(comprador)

        # Preenche a Loja
        self._preencher_loja(loja)

        # Frete e prazo de pagamento
        self._configurar_frete_e_prazo()

        # Abre a Janela de Digitar Produtos
        self._abrir_janela_digitar_produtos(data_da_proxima_visita)

    def _preencher_comprador(self, comprador):
        """Preenche o campo do comprador na interface."""
        comprador_valido = comprador.lower()
        if comprador_valido == "vago":
            pg.press("up", presses=10); sleep(1)
            pg.press('enter')
        elif comprador_valido == "jaidson":
            pg.press("up", presses=10); sleep(1)
            pg.press("down", presses=2); sleep(1)
            pg.press('enter')
        else:
            print(f"Comprador inválido: {comprador}. A digitação será encerrada.")
            sys.exit()

    def _preencher_loja(self, loja):
        """Preenche o campo da loja na interface."""
        loja = loja.lower()
        opcoes_loja = {"smj": 6, "stt": 7, "vix": 8, "mcp": 9}
        if loja in opcoes_loja:
            pg.click(665, 67, duration=1)
            pg.press('up', presses=6)  # Reset position
            pg.press('down', presses=opcoes_loja[loja] - 6)
            pg.press('enter')
        else:
            print(f"Loja inválida: {loja}. A digitação será encerrada.")
            sys.exit()

    def _configurar_frete_e_prazo(self):
        """Configura o frete e o prazo de pagamento na interface."""
        pg.hotkey("alt", "f")
        pg.click(205, 240, duration=1)
        pg.hotkey('alt', 'i')
        pg.click(373, 187, duration=1)
        pg.typewrite('28')  # Exemplo de valor para o prazo
        pg.press('enter')

    def _abrir_janela_digitar_produtos(self, data_da_proxima_visita):
        """Abre a janela para digitar os produtos e lida com possíveis alertas de data da próxima visita."""
        # Aguarda a janela de digitar produtos abrir
        pg.hotkey('alt', 'p')
        sleep(1)
        pg.click(35, 933, duration=1)
        sleep(2)

        proxima_visita_img = self.util_imagem.criar_caminho_imagem("data_proxima_visita", self.configs.caminho_imagens)
        if pg.locateOnScreen(proxima_visita_img) is not None:
            self._corrigir_data_proxima_visita(data_da_proxima_visita)
        else:
            self.util_imagem.sleep_por_imagem("janela_digitar_produtos.png")

    def _corrigir_data_proxima_visita(self, data_da_proxima_visita):
        """Corrige a data da próxima visita se necessário."""
        pg.press("enter")
        pg.typewrite(data_da_proxima_visita)
        sleep(2)
        pg.hotkey("alt", "p")
        sleep(1)
        pg.click(35, 933, duration=1)
        sleep(2)
        self.util_imagem.sleep_por_imagem("janela_digitar_produtos.png")

class DigitacaoProdutos:
    def __init__(self, util_imagem, configs):
        self.util_imagem = util_imagem
        self.configs = configs

    def digitar_produtos(self, tipo_pedido, arquivo_completo):
        with open(arquivo_completo) as arquivo:
            linhas_arquivo = arquivo.readlines()

            for linha in linhas_arquivo:
                if tipo_pedido == "cotação":
                    codigo, embalagem, quantidade, preco = linha.strip().split(";")
                elif tipo_pedido == "compras":
                    codigo, quantidade, embalagem = linha.strip().split(";")

                pg.write(codigo)
                pg.press("enter")
                sleep(0.5)

                if self.verificar_erros_produto(codigo):
                    continue

                pg.press("enter")
                pg.write(quantidade)
                pg.press("enter")

                if tipo_pedido == "cotação":
                    pg.doubleClick(343, 574, duration=0.5)
                    pg.write(preco)
                    pg.press("enter")

                pg.hotkey("alt", "o")
                sleep(1)

                if self.verificar_erros_produto(codigo, erro_inicial=False):
                    continue

            # Fecha a janela de digitar produtos
            pg.press("esc")
            sleep(3)

    def verificar_erros_produto(self, codigo, erro_inicial=True):
        bloqueado_img = self.util_imagem.criar_caminho_imagem("bloqueado", self.configs.caminho_imagens)
        fora_do_mix_img = self.util_imagem.criar_caminho_imagem("fora_do_mix", self.configs.caminho_imagens)
        produto_ja_existe_img = self.util_imagem.criar_caminho_imagem("produto_ja_existe", self.configs.caminho_imagens)
        preco_de_venda_img = self.util_imagem.criar_caminho_imagem("venda", self.configs.caminho_imagens)
        custo_unitario_img = self.util_imagem.criar_caminho_imagem("unitario", self.configs.caminho_imagens)
        
        erro_encontrado = False
        
        if erro_inicial:
            if pg.locateOnScreen(bloqueado_img) is not None:
                logging.warning(f"Produto {codigo} bloqueado para pedido de compra.")
                erro_encontrado = True
            elif pg.locateOnScreen(fora_do_mix_img) is not None:
                logging.warning(f"Produto {codigo} fora do mix.")
                erro_encontrado = True
            elif pg.locateOnScreen(produto_ja_existe_img) is not None:
                logging.warning(f"Produto {codigo} já existe no pedido.")
                erro_encontrado = True
        else:
            if pg.locateOnScreen(preco_de_venda_img) is not None:
                logging.warning(f"Produto {codigo} sem preço de venda.")
                erro_encontrado = True
            elif pg.locateOnScreen(custo_unitario_img) is not None:
                logging.warning(f"Produto {codigo} sem custo unitário.")
                erro_encontrado = True
        
        if erro_encontrado:
            self.produto_bloqueado_ou_fora_do_mix() if erro_inicial else self.sem_preco_de_venda_ou_custo_unitario()
            return True
        else:
            return False

    def produto_bloqueado_ou_fora_do_mix(self):
        # Implementação da lógica para lidar com produtos bloqueados ou fora do mix
        pg.press("enter"); sleep(1)
        pg.press("tab", presses=5)

    def sem_preco_de_venda_ou_custo_unitario(self):
        # Implementação da lógica para lidar com produtos sem preço de venda ou custo unitário
        pg.press("enter"); sleep(2)
        pg.doubleClick(344, 573, duration=0.5); sleep(1)
        pg.typewrite("99")
        pg.press("enter"); sleep(1)
        pg.hotkey("alt", "o")

class SalvamentoPedido:
    def __init__(self, util_imagem, configs):
        self.util_imagem = util_imagem
        self.configs = configs

    def salvar_pedido(self, fornecedor, loja, num_do_pedido, data_pedido, tipo_pedido, valor_formatado, data_historico):
        pg.press("f11")
        self.util_imagem.sleep_por_imagem(os.path.join(self.configs.caminho_imagens, "salvar_pedido.png"))
        
        pg.hotkey("alt", "o")

        self.util_imagem.sleep_por_imagem(os.path.join(self.configs.caminho_imagens, "salvar_pdf.PNG"))
        sleep(2)
        
        pg.click(465, 39, duration=0.5)
        sleep(2)

        if self.desativar_capslock():
            pg.press('capslock')
        
        dados = f"{fornecedor} {loja} {num_do_pedido} {data_pedido}"
        
        logging.info(f"{dados}.pdf salvo com sucesso!")
        
        arquivo_com_extensao = f"{dados}.pdf"
        arquivo_downloads = os.path.join(os.environ.get("USERPROFILE"), "Downloads", arquivo_com_extensao)
        
        pg.typewrite(dados)
        sleep(1)
        
        pg.press("enter")
        sleep(1)
        
        pg.press("f9")
        sleep(2)

        pg.press("esc")
        sleep(1)

        pg.press("f9", presses=2, interval=2)

        if tipo_pedido == "compras":
            self.registrar_historico_compras([num_do_pedido, fornecedor, loja, valor_formatado, data_historico])
        
        return arquivo_downloads

    def desativar_capslock(self):
        """Verifica se o Caps Lock está ativo e retorna True se estiver, False caso contrário."""
        return ctypes.windll.user32.GetKeyState(0x14) & 1 != 0

    def registrar_historico_compras(self, dados_lista):
        """Registra o pedido no arquivo histórico de compras."""
        caminho_arquivo = os.path.join("Outros", "Historico_compras.csv")
        with open(caminho_arquivo, 'a', newline='', encoding='utf-8') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            arquivo_vazio = arquivo.tell() == 0
            if arquivo_vazio:
                escritor.writerow(['Numero do Pedido', 'Fornecedor', 'Loja', 'Valor do Pedido', 'Data'])
            escritor.writerow(dados_lista)