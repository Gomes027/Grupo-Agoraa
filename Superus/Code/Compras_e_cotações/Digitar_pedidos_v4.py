import os
import csv
import sys
import ctypes
import shutil
import locale
import chardet
import logging
import getpass
import keyboard
import subprocess
import pytesseract
import pyautogui as pg
from io import BytesIO
from time import sleep
from PIL import Image, ImageEnhance
from datetime import date, timedelta
from Compras_e_cotações.config import *


class ConfiguracoesIniciais:
    def __init__(self):
        """Inicializa as configurações iniciais."""
        self.usuario_atual = getpass.getuser()
        self.patch_tesseract = f'C:\\Users\\{self.usuario_atual}\\AppData\\Local\\Programs\\Tesseract-OCR\\tesseract.exe'
        pytesseract.pytesseract.tesseract_cmd = self.patch_tesseract
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


class UtilitariosDeImagem:
    @staticmethod
    def criar_caminho_imagem(nome_imagem, caminho_base):
        """
        Cria o caminho completo para um arquivo de imagem.

        Args:
            nome_imagem (str): Nome do arquivo de imagem.
            caminho_base (str): Caminho base onde a imagem está localizada.

        Returns:
            str: Caminho completo para a imagem.
        """
        return os.path.join(caminho_base, f"{nome_imagem}.png")

    @staticmethod
    def capturar_numero_do_pedido(region):
        """
        Captura o número do pedido de uma região específica da tela.

        Args:
            region (tuple): Uma tupla (x, y, largura, altura) que define a região da tela a ser capturada.

        Returns:
            str: O número do pedido como string.
        """
        imagem = UtilitariosDeImagem.capturar_e_processar_imagem(region)
        return pytesseract.image_to_string(imagem, config='--psm 6', lang='por').strip().replace('\n', '').replace('\r', '')

    @staticmethod
    def capturar_e_processar_imagem(region):
        """
        Captura e processa uma região específica da tela.

        Args:
            region (tuple): Uma tupla (x, y, largura, altura) que define a região da tela a ser capturada.

        Returns:
            Image: A imagem processada.
        """
        screenshot = pg.screenshot(region=region)
        imagem_bytesio = BytesIO()
        screenshot.save(imagem_bytesio, format='PNG')
        imagem = Image.open(imagem_bytesio)
        imagem = imagem.resize((imagem.width * 2, imagem.height * 2), Image.LANCZOS)
        imagem = ImageEnhance.Contrast(imagem).enhance(2)
        return imagem
    
    @staticmethod
    def procurar_botao_e_clicar(imagem, clicar=True):
        """
        Procura um botão na tela e realiza um clique se encontrado.

        Args:
            imagem (str): Nome do arquivo de imagem representando o botão.
            clicar (bool): Indica se deve ser realizado o clique no botão encontrado.
        """

        if imagem == "janela_digitar_produtos.png":
            valor_conf = 0.8
        else:
            valor_conf = 0.9

        while True:
            try:
                caminho_imagem = os.path.join("Imgs", imagem)
                localizacao = pg.locateOnScreen(caminho_imagem, confidence=valor_conf)
                
                if localizacao:
                    if clicar:
                        pg.click(pg.center(localizacao), duration=0.1)
                    break
            except Exception:
                continue
    
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


class GerenciamentoDeArquivos:
    def __init__(self):
        """Inicializa o gerenciamento de arquivos."""
        self.usuario_atual = getpass.getuser()

    @staticmethod
    def extrair_fornecedor_loja_comprador(arquivo_completo):
        """
        Extrai o fornecedor, a loja e o comprador de um arquivo.

        :param arquivo_completo: O caminho completo do arquivo a ser processado.
        :return: Uma tupla contendo o fornecedor, a loja e o comprador (se encontrados) ou None, None, None se não forem encontrados.
        """
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
        """
        Remove produtos do arquivo que não possuem quantidade especificada.

        :param arquivo_completo: O caminho completo do arquivo a ser processado.
        """
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
    def __init__(self):
        """Inicializa o gerenciamento do Superus."""
        self.usuario_atual = getpass.getuser()

    def iniciar_superus(self):
        """Inicia o aplicativo Superus."""
        subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
        self.aguardar_abertura()

    def aguardar_abertura(self):
        """Aguarda a abertura do aplicativo Superus e realiza ações específicas dependendo do usuário."""
        UtilitariosDeImagem.procurar_botao_e_clicar("superus.png", False)

        if self.usuario_atual == "automacao.compras":
            pg.click(572, 574, duration=0.5); sleep(1)
            pg.write("123456"); sleep(3)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(615, 512, duration=0.5); sleep(1)
            pg.write("848600"); sleep(3)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        pg.press("enter", presses=2, interval=1)
        UtilitariosDeImagem.procurar_botao_e_clicar("superus_aberto.png", False)

    @staticmethod
    def fechar_processo():
        """Fecha o processo do aplicativo Superus."""
        try:
            subprocess.run(['taskkill', '/f', '/im', "SUPERUS.EXE"], check=True)
        except subprocess.CalledProcessError:
            pass


class OperacoesDePedido:
    def __init__(self):
        """Inicializa as instâncias necessárias para as operações.

        Args:
            util_imagem (objeto): Objeto utilizado para manipulação de imagens.
            configs (objeto): Objeto contendo as configurações iniciais.
            preenchimento_info (objeto): Objeto para preenchimento de informações de pedido.
            digitar_ped (objeto): Objeto para digitação de produtos.
            salvar_ped (objeto): Objeto para salvamento de pedidos.
            ger_arquivos (objeto): Objeto para gerenciamento de arquivos.
        """
        self.usuario_atual = getpass.getuser()
        self.configs = ConfiguracoesIniciais()
        self.util_imagem = UtilitariosDeImagem()
        self.ger_arquivos = GerenciamentoDeArquivos()
        
        # Passando as instâncias necessárias como argumentos para os construtores
        self.preenchimento_info = PreenchimentoPedido(self.util_imagem, self.configs)
        self.digitar_ped = DigitacaoProdutos(self.util_imagem, self.configs)
        self.salvar_ped = SalvamentoPedido(self.util_imagem, self.configs)

    def executar_automacao(self, arquivo_completo, tipo_pedido, novo_arquivo):
        """Executa a automação para processar um pedido.

        Args:
            arquivo_completo (str): Nome completo do arquivo.
            tipo_pedido (str): Tipo de pedido.
            novo_arquivo (str): Nome do novo arquivo.
        """
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
        """Insere o controle de pedidos na interface do sistema."""
        if self.usuario_atual == "automacao.compras":
            pg.doubleClick(381, 206, duration=0.5)
        elif self.usuario_atual == "automacao.compras1":
            pg.doubleClick(452, 140, duration=0.5)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")
        
        sleep(5)
        pg.write("81569"); pg.press("enter"); sleep(5)
        
        preencher_pedido = False
        
        try:
            if pg.locateOnScreen(self.util_imagem.criar_caminho_imagem("produto_ja_existe", "Imgs"), confidence=0.9) is not None:
                self.digitar_ped.produto_bloqueado_ou_fora_do_mix()
            else:
                preencher_pedido = True
        except Exception:
            preencher_pedido = True
        
        if preencher_pedido:
            if self.usuario_atual == "automacao.compras":
                pg.doubleClick(734, 389, duration=0.5); pg.write("1")
                pg.doubleClick(343, 574, duration=0.5); pg.write("0,01")
            elif self.usuario_atual == "automacao.compras1":
                pg.doubleClick(821, 326, duration=0.5); pg.write("1")
                pg.doubleClick(426, 512, duration=0.5); pg.write("0,01")
            else:
                logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

            pg.press("enter"); pg.hotkey("alt", "o"); sleep(3)

    def atualizar_data(self):
        """
        Atualiza as datas necessárias para o processamento do pedido.

        Returns:
            tuple: Uma tupla contendo a data formatada, o formato do pedido e a data de correção.
        """
        data_atual = date.today()
        data_formatada = data_atual.strftime("%d/%m/%Y")
        data_pedido = data_atual.strftime("%d.%m")
        data_correcao = (data_atual + timedelta(days=7)).strftime("%d%m%Y")
        
        return data_formatada, data_pedido, data_correcao

    def obter_numero_do_pedido(self):
        """
        Obtém o número do pedido a partir de uma região específica na tela.

        Returns:
            int: O número do pedido obtido.
        """
        region = (73, 63, 75, 25)
        num_do_pedido = self.util_imagem.capturar_numero_do_pedido(region)
        try:
            num_do_pedido = int(num_do_pedido.strip().replace('\n', '').replace('\r', ''))
        except ValueError:
            print("Erro ao converter o número do pedido para inteiro.")
            num_do_pedido = 0
        
        return num_do_pedido

    def obter_valor_do_pedido(self):
        """
        Obtém o valor do pedido e formata para exibição.

        Returns:
            str: O valor do pedido formatado para exibição.
        """
        region = (305, 126, 75, 25)
        valor_do_pedido = self.util_imagem.capturar_valor_do_pedido(region)
        try:
            valor_do_pedido = float(valor_do_pedido.strip().replace(',', '.'))
            valor_formatado = f"R$ {valor_do_pedido:,.2f}".replace(',', ' temporário ').replace('.', ',').replace(' temporário ', '.')
        except ValueError:
            print("Erro ao converter o valor do pedido.")
            valor_formatado = "R$ 0,00"
        
        return valor_formatado

    def mover_arquivo_pedido(self, arquivo_downloads, tipo_pedido):
        """Move o arquivo de pedido para a pasta correta.

        Args:
            arquivo_downloads (str): Caminho completo do arquivo a ser movido.
            tipo_pedido (str): Tipo de pedido, pode ser "compras" ou "cotação".
        """
        
        if tipo_pedido == "compras":
            destino = DIR_PDF_COMPRAS
        elif tipo_pedido == "cotação":
            destino = DIR_PDF_COTACAO
        else:
            print(f"Tipo de pedido desconhecido: {tipo_pedido}")
            return
        
        shutil.move(arquivo_downloads, destino)


class PreenchimentoPedido:
    def __init__(self, util_imagem, configs):
        """Inicializa a classe com as instâncias necessárias.

        Args:
            util_imagem (objeto): Objeto utilizado para manipulação de imagens.
            configs (objeto): Objeto contendo as configurações necessárias.
        """
        self.usuario_atual = getpass.getuser()
        self.util_imagem = util_imagem
        self.configs = configs

    def preencher_informacoes_pedido(self, fornecedor, loja, comprador, data_da_proxima_visita, arquivo_completo, tipo_pedido):
        """Preenche as principais informações do pedido na interface.

        Args:
            fornecedor (str): Nome do fornecedor.
            loja (str): Nome da loja.
            comprador (str): Nome do comprador.
            data_da_proxima_visita (str): Data da próxima visita.
            arquivo_completo (str): Nome do arquivo completo.
            tipo_pedido (str): Tipo do pedido.
        """
        self.util_imagem.procurar_botao_e_clicar(r"compras.png")
        self.util_imagem.procurar_botao_e_clicar(r"pedidos_de_compra.png"); sleep(6)
        pg.press('F2')
        
        self.util_imagem.procurar_botao_e_clicar(r"pedido_compra.png", False)

        if self.usuario_atual == "automacao.compras":
            pg.click(143, 100, duration=1)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(135, 99, duration=1)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        sleep(2); keyboard.write(fornecedor)
        pg.press('enter'); pg.hotkey('alt', 'o'); sleep(4)
        
        try:
            if pg.locateOnScreen(r"Imgs\fornecedor_recadastramento.png", confidence=0.8) is not None:
                pg.press("tab"); pg.press("enter")
        except Exception:
            pass
        
        self._preencher_comprador(comprador)
        self._preencher_loja(loja)
        self._configurar_frete_e_prazo()
        self._abrir_janela_digitar_produtos(data_da_proxima_visita)

    def _preencher_comprador(self, comprador):
        """Preenche o campo do comprador na interface.

        Args:
            comprador (str): Nome do comprador.
        """
        if self.usuario_atual == "automacao.compras":
            pg.click(669, 87, duration=3)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(670, 91, duration=3)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        sleep(1); comprador_valido = comprador.lower()
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
        """Preenche o campo da loja na interface.

        Args:
            loja (str): Nome da loja.
        """
        if self.usuario_atual == "automacao.compras":
            pg.click(665, 67, duration=1)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(669, 68, duration=1)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        loja = loja.lower()
        opcoes_loja = {"smj": 6, "stt": 7, "vix": 8, "mcp": 9}
        if loja in opcoes_loja:
            pg.press('up', presses=6)  # Reset position
            pg.press('down', presses=opcoes_loja[loja] - 6)
            pg.press('enter')
        else:
            print(f"Loja inválida: {loja}. A digitação será encerrada.")
            sys.exit()

    def _configurar_frete_e_prazo(self):
        """Configura o frete e o prazo de pagamento na interface."""
        pg.hotkey("alt", "f"); pg.click(205, 240, duration=1)
        pg.hotkey('alt', 'i'); pg.click(373, 187, duration=1)
        pg.typewrite('28'); pg.press('enter')

    def _abrir_janela_digitar_produtos(self, data_da_proxima_visita):
        """Abre a janela para digitar os produtos e lida com possíveis alertas de data da próxima visita.

        Args:
            data_da_proxima_visita (str): Data da próxima visita.
        """
        pg.hotkey('alt', 'p'); sleep(1)

        if self.usuario_atual == "automacao.compras":
            pg.click(35, 933, duration=1)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(32, 814, duration=1)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        try:
            if pg.locateOnScreen(r"Imgs\data_proxima_visita.png") is not None:
                pg.press("enter"); self._corrigir_data_proxima_visita(data_da_proxima_visita)
        except Exception:
                self.util_imagem.procurar_botao_e_clicar("janela_digitar_produtos.png", False)

    def _corrigir_data_proxima_visita(self, data_da_proxima_visita):
        """Corrige a data da próxima visita se necessário.

        Args:
            data_da_proxima_visita (str): Data da próxima visita.
        """
        pg.typewrite(data_da_proxima_visita); sleep(2)
        pg.hotkey("alt", "p"); sleep(1)

        if self.usuario_atual == "automacao.compras":
            pg.click(35, 933, duration=1)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(32, 814, duration=1)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        self.util_imagem.procurar_botao_e_clicar("janela_digitar_produtos.png", False)


class DigitacaoProdutos:
    def __init__(self, util_imagem, configs):
        """Inicializa a classe com as instâncias necessárias.

        Args:
            util_imagem (objeto): Instância da classe UtilImagem.
            configs (dict): Dicionário contendo as configurações necessárias.
        """
        self.usuario_atual = getpass.getuser()
        self.util_imagem = util_imagem
        self.configs = configs

    def digitar_produtos(self, tipo_pedido, arquivo_completo):
        """Digita os produtos conforme o tipo de pedido.

        Args:
            tipo_pedido (str): Tipo de pedido a ser processado ("cotação" ou "compras").
            arquivo_completo (str): Caminho completo para o arquivo a ser processado.
        """
        with open(arquivo_completo) as arquivo:
            linhas_arquivo = arquivo.readlines()

            for linha in linhas_arquivo:
                if tipo_pedido == "cotação":
                    codigo, embalagem, quantidade, preco = linha.strip().split(";")
                elif tipo_pedido == "compras":
                    codigo, quantidade, embalagem = linha.strip().split(";")

                pg.write(codigo); pg.press("enter"); sleep(1)

                if self.verificar_erros_produto(codigo):
                    continue

                pg.press("enter")
                pg.write(quantidade)
                pg.press("enter")

                if tipo_pedido == "cotação":
                    if self.usuario_atual == "automacao.compras":
                        pg.doubleClick(343, 574, duration=0.5)
                    elif self.usuario_atual == "automacao.compras1":
                        pg.doubleClick(426, 512, duration=0.5)
                    else:
                        logging.error(f"Usuário {self.usuario_atual} não reconhecido.")
                    
                    pg.write(preco); pg.press("enter")

                pg.hotkey("alt", "o"); sleep(2)

                if self.verificar_erros_produto(codigo, erro_inicial=False):
                    continue

            pg.press("esc")
            sleep(3); pg.press("f11")

    def verificar_erros_produto(self, codigo, erro_inicial=True):
        """Verifica se há erros relacionados ao produto.

        Args:
            codigo (str): Código do produto a ser verificado.
            erro_inicial (bool, optional): Indica se é a primeira verificação de erro. 
                                           Defaults to True.

        Returns:
            bool: True se um erro for encontrado, False caso contrário.
        """
        bloqueado_img = self.util_imagem.criar_caminho_imagem("bloqueado", "Imgs")
        fora_do_mix_img = self.util_imagem.criar_caminho_imagem("fora_do_mix", "Imgs")
        produto_ja_existe_img = self.util_imagem.criar_caminho_imagem("produto_ja_existe", "Imgs")
        preco_de_venda_img = self.util_imagem.criar_caminho_imagem("venda", "Imgs")
        custo_unitario_img = self.util_imagem.criar_caminho_imagem("unitario", "Imgs")
        
        verificacoes = [
            (bloqueado_img, f"Produto {codigo} bloqueado para pedido de compra."),
            (fora_do_mix_img, f"Produto {codigo} fora do mix."),
            (produto_ja_existe_img, f"Produto {codigo} já existe no pedido."),
            (preco_de_venda_img, f"Produto {codigo} sem preço de venda."),
            (custo_unitario_img, f"Produto {codigo} sem custo unitário.")
        ]

        erro_encontrado = False

        if self.usuario_atual == "automacao.compras":
            valor_confidence = 1
        elif self.usuario_atual == "automacao.compras1":
            valor_confidence = 0.9
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        for img, mensagem in verificacoes:
            if not erro_encontrado:
                try:
                    if pg.locateOnScreen(img, confidence=valor_confidence) is not None:
                        logging.warning(mensagem)
                        erro_encontrado = True
                except Exception:
                    continue
        
        if erro_encontrado:
            self.produto_bloqueado_ou_fora_do_mix() if erro_inicial else self.sem_preco_de_venda_ou_custo_unitario()
            return True
        else:
            return False

    def produto_bloqueado_ou_fora_do_mix(self):
        """Implementação da lógica para lidar com produtos bloqueados ou fora do mix."""
        pg.press("enter"); sleep(2)
        pg.press("tab", presses=5)

    def sem_preco_de_venda_ou_custo_unitario(self):
        """Implementação da lógica para lidar com produtos sem preço de venda ou custo unitário."""
        pg.press("enter"); sleep(2)

        if self.usuario_atual == "automacao.compras":
            pg.doubleClick(344, 573, duration=0.5)
        elif self.usuario_atual == "automacao.compras1":
            pg.doubleClick(426, 512, duration=0.5)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        sleep(1); pg.typewrite("99")
        pg.press("enter"); sleep(1); pg.hotkey("alt", "o"); sleep(1)


class SalvamentoPedido:
    def __init__(self, util_imagem, configs):
        """Inicializa a classe com as instâncias necessárias.

        Args:
            util_imagem (objeto): Instância da classe UtilImagem.
            configs (dict): Dicionário contendo as configurações necessárias.
        """ 
        self.usuario_atual = getpass.getuser()
        self.util_imagem = util_imagem
        self.configs = configs

    def salvar_pedido(self, fornecedor, loja, num_do_pedido, data_pedido, tipo_pedido, valor_formatado, data_historico):
        """
        Salva o pedido e move o arquivo PDF conforme o tipo de pedido.

        Args:
            fornecedor (str): Nome do fornecedor.
            loja (str): Nome da loja.
            num_do_pedido (int): Número do pedido.
            data_pedido (str): Data do pedido no formato dd.mm.
            tipo_pedido (str): Tipo de pedido ('compras' ou 'cotação').
            valor_formatado (str): Valor do pedido formatado.
            data_historico (str): Data para registro no histórico de compras.

        Returns:
            str: Caminho completo do arquivo de pedido baixado.
        """
        self.util_imagem.procurar_botao_e_clicar("salvar_pedido.png", False); pg.hotkey("alt", "o")
        self.util_imagem.procurar_botao_e_clicar("salvar_pdf.PNG", False); sleep(2)

        if self.usuario_atual == "automacao.compras":
            pg.click(465, 39, duration=0.5)
        elif self.usuario_atual == "automacao.compras1":
            pg.click(463, 43, duration=0.5)
        else:
            logging.error(f"Usuário {self.usuario_atual} não reconhecido.")

        sleep(2)

        if self.desativar_capslock():
            pg.press('capslock')
        
        dados = f"{fornecedor} {loja} {num_do_pedido} {data_pedido}"
        
        logging.info(f"{dados}.pdf salvo com sucesso!")
        
        arquivo_com_extensao = f"{dados}.pdf"
        arquivo_downloads = os.path.join(os.environ.get("USERPROFILE"), "Downloads", arquivo_com_extensao)
        
        pg.typewrite(dados); sleep(1)
        pg.press("enter"); sleep(1)
        pg.press("f9"); sleep(2)
        pg.press("esc"); sleep(1)
        pg.press("f9", presses=2, interval=2)

        if tipo_pedido == "compras":
            self.registrar_historico_compras([data_historico, num_do_pedido, fornecedor, loja, valor_formatado])
        
        return arquivo_downloads

    def desativar_capslock(self):
        """
        Verifica se o Caps Lock está ativo e retorna True se estiver, False caso contrário.

        Returns:
            bool: True se o Caps Lock estiver ativo, False caso contrário.
        """
        return ctypes.windll.user32.GetKeyState(0x14) & 1 != 0

    def registrar_historico_compras(self, dados_lista):
        """
        Registra o pedido no arquivo histórico de compras.

        Args:
            dados_lista (list): Lista contendo os dados do pedido a serem registrados.
        """
        caminho_arquivo = os.path.join("Outros", "Historico_compras.csv")
        with open(caminho_arquivo, 'a', newline='', encoding='utf-8') as arquivo:
            escritor = csv.writer(arquivo, delimiter=';')
            arquivo_vazio = arquivo.tell() == 0
            if arquivo_vazio:
                escritor.writerow(['Data', 'Numero do Pedido', 'Fornecedor', 'Loja', 'Valor do Pedido'])
            escritor.writerow(dados_lista)