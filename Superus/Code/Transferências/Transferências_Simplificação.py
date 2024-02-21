import os
import time
import logging
import datetime
import subprocess
from time import sleep
import pyautogui as pg

class TransferenciasEntreLojas:
    data_hoje = datetime.date.today().strftime('%d-%m-%Y')

    # Configura o nome do arquivo de log para incluir a data de hoje
    nome_arquivo_log = f"Logs\\Transfêrencias\\Transferencias_{data_hoje}.log"

    # Configuração do logging
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S',
                        filename=nome_arquivo_log,  # Usa o nome do arquivo com a data
                        filemode='a')

    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)
    
    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

    def fechar_processo(self, nome_do_processo):
        """Fecha o processo especificado."""
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
            logging.info("Processo %s fechado com sucesso.", nome_do_processo)
        except subprocess.CalledProcessError:
            logging.critical("Não foi possível fechar o processo %s.", nome_do_processo)

    def procurar_botao_e_clicar(self, imagem, clicar=True):
        """Procura por um botão baseado na imagem e clica se especificado."""
        while True:
            localizacao = pg.locateOnScreen(os.path.join("Imgs", imagem))
            if localizacao:
                if clicar:
                    pg.click(pg.center(localizacao), duration=0.5)
                break

    def login_simplificacao(self):
        """Inicia o Simplificação e navega pelos menus iniciais."""
        logging.info("Iniciando simplificação.")
        os.startfile(r"C:\SimplificAcao_Valid\SmartApp.exe")
        for imagem in ["logar.png", "operacoes.png", "loja_loja.png"]:
            self.procurar_botao_e_clicar(imagem)

    def selecionar_lojas(self):
        """Seleciona as lojas com base no nome do arquivo."""
        partes = self.nome_arquivo.split()
        if len(partes) >= 2:
            loja_1, loja_2 = partes[0], partes[2]

        loja_que_pede = f"{loja_2}.png"
        loja_que_responde = f"{loja_2}/{loja_1}.png"

        # Adicionando logging para loja que pede e loja que responde
        logging.info("Transfêrencia de para: %s -> %s", loja_1, loja_2)

        for imagem in [loja_que_pede, loja_que_responde, "digitar.png"]:
            self.procurar_botao_e_clicar(imagem); sleep(3)

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
        logging.info("O arquivo %s contém %d linhas.", self.nome_arquivo, quantidade_linhas)

        logging.info("Iniciando processamento do arquivo %s.", self.nome_arquivo)

        for linha in linhas:
            try:
                codigo, quantidade = linha.strip().split(";")
                pg.typewrite(codigo); pg.press("enter")
                self.procurar_botao_e_clicar("qtde_2.png", False)
                pg.typewrite(quantidade); sleep(1); pg.press("enter")
                self.procurar_botao_e_clicar("menu_digitacao.png", False); sleep(1)
                self.remover_linha(linhas, linha)
                logging.info("Código %s processado, %s unidades", codigo, quantidade)
            except Exception as e:
                print(e)
        
        logging.info("Enviando arquivo %s...", self.nome_arquivo)
        self.procurar_botao_e_clicar(r"enviar_arquivo.png")
        self.procurar_botao_e_clicar(r"confirmar.png", clicar=False); pg.press("enter")
        self.procurar_botao_e_clicar(r"ok.png", clicar=False); pg.press("enter")
        logging.info("Arquivo %s, enviado com sucesso!", self.nome_arquivo)

        fim_processamento_arquivo = time.time()  # Marca o fim do processamento do arquivo
        duracao_total = fim_processamento_arquivo - inicio_processamento_arquivo  # Calcula a duração total do processamento
        
        # Verifica se a duração é menor que 60 segundos
        if duracao_total < 60:
            tempo_formatado = "%.2f segundos" % duracao_total
        else:
            minutos = duracao_total / 60
            tempo_formatado = "%.2f minutos" % minutos

        logging.info("Processamento do arquivo %s concluído em %s.", self.nome_arquivo, tempo_formatado)

    def executar(self):
        """Orquestra o processo completo."""
        self.login_simplificacao()
        self.selecionar_lojas()
        self.processar_arquivo()
        self.fechar_processo("SmartApp.exe")