import os
import csv
import chardet
import logging
import datetime
from time import sleep
from filelock import FileLock, Timeout

# Importações dos módulos externos
from Validades.src.Calculadora_de_validades import processar_arquivo
from Compras_e_cotações.Enviar_Pedidos_WhatsApp import AutomacaoWhatsApp
from Transferências.Transferências_Simplificação import TransferenciasEntreLojas
from Compras_e_cotações.Digitar_pedidos_v3_0 import OperacoesDePedido, GerenciamentoSuperus

class GerenciadorDeLog:
    @staticmethod
    def configurar_logging():
        nome_usuario = os.getlogin()
        data_hoje = datetime.date.today()
        nome_arquivo_log_formatado = f"Logs\Log_{nome_usuario}_{data_hoje}.log"

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

class GerenciadorDePedidos:
    DIR_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos"
    DIR_PDF_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"
    DIR_PDF_COTACAO = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"

    def __init__(self, automacao_whatsapp=None):
        self.gerenciador_de_logging = GerenciadorDeLog()
        self.gerenciador_de_logging.configurar_logging()
        self.gerenciador_superus = GerenciamentoSuperus()
        self.superus_iniciado = False
        self.automacao_whatsapp = automacao_whatsapp

    def verificar_tipo_arquivo(arquivo_completo):
        try:
            with open(arquivo_completo, 'rb') as file:
                encoding = chardet.detect(file.read())['encoding']

            with open(arquivo_completo, newline='', encoding=encoding) as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                linhas = list(reader)
                if not linhas:
                    return "arquivo vazio"

                primeira_linha = linhas[0]
                if len(linhas) == 3 and not any(char.isdigit() for char in primeira_linha[0]):
                    return "relatorio_pedidos"

                if len(primeira_linha) == 2:
                    return "transferência"
                elif len(primeira_linha) == 3:
                    return "compras"
                elif len(primeira_linha) == 4:
                    return "cotação"
                else:
                    return "formato desconhecido"
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
            return "erro de leitura"

    def pontuacao_tipo_arquivo(self, tipo_pedido):
        return {"transferência": 1, "cotação": 2, "compras": 3}.get(tipo_pedido, 4)
    
    def processar_validades(self, validades):
        for validade in validades:
            caminho_validade = os.path.join(self.DIR_PEDIDOS, validade)
            lock = FileLock(f"{caminho_validade}.lock")
            try:
                with lock.acquire(timeout=0.1):  # Tenta adquirir o bloqueio, esperando por 0.1 segundo
                    processar_arquivo(caminho_validade)
                os.remove(caminho_validade)
            except (Timeout, FileNotFoundError):
                continue

    def processar(self):
        while True:
            validades = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".xlsx")}
            arquivos = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".csv")}
            pdfs_cotacao = {arq for arq in os.listdir(self.DIR_PDF_COTACAO) if arq.lower().endswith(".pdf")}
            pdfs_compras = {arq for arq in os.listdir(self.DIR_PDF_PEDIDOS) if arq.lower().endswith(".pdf")}

            if validades:
                self.processar_validades(validades)

            arquivos_ordenados = sorted(
                arquivos,
                key=lambda arq: (self.pontuacao_tipo_arquivo(self.verificar_tipo_arquivo(os.path.join(self.DIR_PEDIDOS, arq))), arq)
            )

            for arquivo in arquivos_ordenados:
                arquivo_completo = os.path.join(self.DIR_PEDIDOS, arquivo)
                lock = FileLock(f"{arquivo_completo}.lock")

                try:
                    with lock.acquire(timeout=0.1):  # Tenta adquirir o bloqueio, esperando por 0.1 segundo
                        tipo_pedido = self.verificar_tipo_arquivo(arquivo_completo)

                        if tipo_pedido == "transferência":
                            TransferenciasEntreLojas(arquivo_completo).executar()
                        elif tipo_pedido in ["cotação", "compras"]:
                            if not self.superus_iniciado:
                                self.gerenciador_superus.iniciar_superus()
                                self.superus_iniciado = True

                            OperacoesDePedido().executar_automacao(arquivo_completo, tipo_pedido, arquivo)

                        os.remove(arquivo_completo)
                except Timeout:
                    continue  # Se o arquivo estiver bloqueado, continue para o próximo
                
                arquivos = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".csv")}
                validades = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".xlsx")}

                if validades:
                    self.processar_validades(validades)

            if self.superus_iniciado:
                self.gerenciador_superus.fechar_processo()
                self.superus_iniciado = False

            if pdfs_cotacao:
                automacao_whatsapp.processar_pdfs(pdfs_cotacao, "cotação")

            elif pdfs_compras:
                automacao_whatsapp.processar_pdfs(pdfs_compras, "compras")

            sleep(5)

if __name__ == "__main__":
    os.system("cls")
    automacao_whatsapp = AutomacaoWhatsApp()
    gerenciador = GerenciadorDePedidos(automacao_whatsapp)
    gerenciador.processar()