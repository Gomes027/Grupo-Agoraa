import os
import time
import logging
import subprocess
from time import sleep
import pyautogui as pg

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
            try:
                caminho_imagem = os.path.join("Imgs", imagem)
                localizacao = pg.locateOnScreen(caminho_imagem, confidence=0.9)
                
                if localizacao:
                    if clicar:
                        pg.click(pg.center(localizacao), duration=0.1)
                    break
            except Exception:
                continue

    @staticmethod
    def proxima_iteracao(imagens):
        while True:
            try:
                for imagem in imagens:
                    # Definindo a confiança baseada no nome da imagem
                    confidence = 0.6 if imagem == 'qtde_2.png' else 0.7
                    localizacao = pg.locateOnScreen(os.path.join("Imgs", imagem), confidence=confidence)
                    if localizacao:
                        return imagem
            except Exception:
                continue

class TransferenciasEntreLojas:
    def __init__(self, caminho_arquivo):
        self.caminho_arquivo = caminho_arquivo
        self.nome_arquivo = os.path.splitext(os.path.basename(caminho_arquivo))[0]

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
        logging.info("Processando %s, com %d produtos.", self.nome_arquivo, quantidade_linhas)

        for linha in linhas[:]:
            if linha.strip():
                codigo, quantidade = linha.strip().split(";")
                pg.typewrite(codigo); pg.press("enter")
                imagem_codigo = AutomacaoGui.proxima_iteracao(["qtde_2.png", "cod_n_encontrado.png", "informe_um_codigo.png"])

                if imagem_codigo == "informe_um_codigo.png":
                    pg.press("enter"); sleep(3)
                    pg.typewrite(codigo); pg.press("enter")
                    logging.warning("Erro ao digitar codigo %s", codigo)
                    imagem_codigo = AutomacaoGui.proxima_iteracao(["qtde_2.png", "cod_n_encontrado.png"])

                if imagem_codigo == "qtde_2.png":
                    pg.typewrite(quantidade); sleep(1); pg.press("enter")
                    imagem_quantidade = AutomacaoGui.proxima_iteracao(["menu_digitacao.png", "erro_ao_gravar_item.png"])
                    if imagem_quantidade == "menu_digitacao.png":
                        pass
                    elif imagem_quantidade == "erro_ao_gravar_item.png":
                        logging.warning("Erro ao gravar codigo %s", codigo)
                        pg.press("enter"); sleep(3)
                        pg.press("tab"); sleep(1)
                        pg.press("enter"); sleep(10)
                        
                    self.remover_linha(linhas, linha)
                    logging.info("Código %s processado, %s unidades", codigo, quantidade)

                sleep(0.5)
        
        AutomacaoGui.procurar_botao_e_clicar(r"enviar_arquivo.png")
        AutomacaoGui.procurar_botao_e_clicar(r"confirmar.png", clicar=False); pg.press("enter")
        AutomacaoGui.procurar_botao_e_clicar(r"ok.png", clicar=False); pg.press("enter")

        fim_processamento_arquivo = time.time()  # Marca o fim do processamento do arquivo
        duracao_total = fim_processamento_arquivo - inicio_processamento_arquivo  # Calcula a duração total do processamento
        
        # Verifica se a duração é menor que 60 segundos
        if duracao_total < 60:
            tempo_formatado = "%d segundos" % duracao_total
        else:
            minutos = int(duracao_total // 60)  # Usa divisão inteira para descartar os segundos
            tempo_formatado = "%d minutos" % minutos

        logging.info("Arquivo processado em %s!", tempo_formatado)

    def executar(self):
        self.login_simplificacao()
        self.selecionar_lojas()
        self.processar_arquivo()
        GerenciamentoDeProcessos.fechar_processo("SmartApp.exe")