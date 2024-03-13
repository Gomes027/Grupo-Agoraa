import os
import time
import logging
import datetime
import subprocess
from time import sleep
import pyautogui as pg

class GerenciadorDeLog:
    @staticmethod
    def configurar_logging(nome_arquivo_log):
        data_hoje = datetime.date.today()
        nome_arquivo_log_formatado = f"Logs\\Transfêrencias\\{nome_arquivo_log}_{data_hoje}.log"

        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s',
                            datefmt='%Y-%m-%d %H:%M',
                            filename=nome_arquivo_log_formatado,
                            filemode='a')

        console = logging.StreamHandler()
        console.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console.setFormatter(formatter)
        logging.getLogger('').addHandler(console)

class GerenciamentoDeProcessos:
    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            logging.critical("Não foi possível fechar o processo %s.", nome_do_processo)

class AutomacaoGui:
    @staticmethod
    def procurar_botao_e_clicar(imagem, clicar=True):
        while True:
            localizacao = pg.locateOnScreen(os.path.join("Imgs", imagem))
            if localizacao:
                if clicar:
                    pg.click(pg.center(localizacao), duration=0.5)
                break

class TransferenciasEntreLojas:
    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

        # Configuração inicial do logging
        GerenciadorDeLog.configurar_logging("Transferencias")

    def login_simplificacao(self):
        # Simplificação da lógica original com uso do método auxiliar
        os.startfile(r"C:\SimplificAcao_Valid\SmartApp.exe")
        for imagem in ["logar.png", "operacoes.png", "loja_loja.png"]:
            AutomacaoGui.procurar_botao_e_clicar(imagem)

    def selecionar_lojas(self):
        """Seleciona as lojas com base no nome do arquivo."""
        partes = self.nome_arquivo.split()
        if len(partes) >= 2:
            loja_1, loja_2 = partes[0], partes[2]

        loja_que_pede = f"{loja_2}.png"
        loja_que_responde = f"{loja_2}/{loja_1}.png"

        for imagem in [loja_que_pede, loja_que_responde, "digitar.png"]:
            AutomacaoGui.procurar_botao_e_clicar(imagem); sleep(3)

    def remover_linha(self, linhas, linha):
        """Remove a linha processada do arquivo."""
        try:
            linhas.remove(linha)
            with open(self.caminho_arquivo, 'w') as arquivo:
                arquivo.writelines(linhas)
        except:
            logging.error("Erro ao remover linha: %s", linha.strip())

    def processar_arquivo(self):
        """Processa cada linha do arquivo especificado."""
        inicio_processamento_arquivo = time.time()  # Marca o início do processamento do arquivo

        with open(self.caminho_arquivo, 'r') as arquivo:
            linhas = arquivo.readlines()

        quantidade_linhas = len(linhas)  # Conta a quantidade de linhas no arquivo
        logging.info("Iniciando processamento de %s, com %d produtos.", self.nome_arquivo, quantidade_linhas)

        for linha in linhas[:]:
            try:
                codigo, quantidade = linha.strip().split(";")
                pg.typewrite(codigo); pg.press("enter")
                AutomacaoGui.procurar_botao_e_clicar("qtde_2.png", False)
                pg.typewrite(quantidade); sleep(1); pg.press("enter")
                AutomacaoGui.procurar_botao_e_clicar("menu_digitacao.png", False); sleep(1)
                self.remover_linha(linhas, linha)
                logging.info("Código %s processado, %s unidades", codigo, quantidade)
            except Exception as e:
                print(e)
        
        AutomacaoGui.procurar_botao_e_clicar(r"enviar_arquivo.png")
        AutomacaoGui.procurar_botao_e_clicar(r"confirmar.png", clicar=False); pg.press("enter")
        AutomacaoGui.procurar_botao_e_clicar(r"ok.png", clicar=False); pg.press("enter")

        fim_processamento_arquivo = time.time()  # Marca o fim do processamento do arquivo
        duracao_total = fim_processamento_arquivo - inicio_processamento_arquivo  # Calcula a duração total do processamento
        
        # Verifica se a duração é menor que 60 segundos
        if duracao_total < 60:
            tempo_formatado = "%.2f segundos" % duracao_total
        else:
            minutos = duracao_total / 60
            tempo_formatado = "%.2f minutos" % minutos

        logging.info("Arquivo processado com sucesso em %s!", tempo_formatado)

    def executar(self):
        self.login_simplificacao()
        self.selecionar_lojas()
        self.processar_arquivo()
        GerenciamentoDeProcessos.fechar_processo("SmartApp.exe")