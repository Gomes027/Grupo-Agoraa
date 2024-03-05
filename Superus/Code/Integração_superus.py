import os
import csv
import chardet
from time import sleep

# Importações dos módulos externos
from Compras_e_cotações.Enviar_Pedidos_WhatsApp import AutomacaoWhatsApp
from Transferências.Transferências_Simplificação import TransferenciasEntreLojas
from Compras_e_cotações.Digitar_pedidos_v3_0 import OperacoesDePedido, GerenciamentoSuperus

class GerenciadorDePedidos:
    DIR_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos"
    DIR_PDF_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"
    DIR_PDF_COTACAO = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"

    def __init__(self):
        self.gerenciador_superus = GerenciamentoSuperus()
        self.superus_iniciado = False

    def verificar_tipo_arquivo(self, arquivo_completo):
        try:
            with open(arquivo_completo, 'rb') as file:
                encoding = chardet.detect(file.read())['encoding']

            with open(arquivo_completo, newline='', encoding=encoding) as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                for linha in reader:
                    if linha:
                        if len(linha) == 2:
                            return "transferência"
                        elif len(linha) == 3:
                            return "compras"
                        elif len(linha) == 4:
                            return "cotação"
                        else:
                            return "formato desconhecido"
            return "arquivo vazio"
        except Exception as e:
            print(f"Erro ao ler arquivo: {e}")
            return "erro de leitura"

    def pontuacao_tipo_arquivo(self, tipo_pedido):
        return {"transferência": 1, "cotação": 2, "compras": 3}.get(tipo_pedido, 4)

    def processar(self):
        while True:
            arquivos = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".csv")}
            pdfs_cotacao = {arq for arq in os.listdir(self.DIR_PDF_COTACAO) if arq.lower().endswith(".pdf")}
            pdfs_compras = {arq for arq in os.listdir(self.DIR_PDF_PEDIDOS) if arq.lower().endswith(".pdf")}

            while arquivos:
                arquivo = sorted(
                    arquivos, 
                    key=lambda arq: (self.pontuacao_tipo_arquivo(self.verificar_tipo_arquivo(os.path.join(self.DIR_PEDIDOS, arq))), arq)
                )[0]
                
                arquivo_completo = os.path.join(self.DIR_PEDIDOS, arquivo)
                tipo_pedido = self.verificar_tipo_arquivo(arquivo_completo)

                if tipo_pedido == "transferência":
                    TransferenciasEntreLojas(arquivo_completo).executar()
                elif tipo_pedido in ["cotação", "compras"]:
                    if not self.superus_iniciado:
                        self.gerenciador_superus.iniciar_superus()
                        self.superus_iniciado = True

                    OperacoesDePedido().executar_automacao(arquivo_completo, tipo_pedido, arquivo)

                os.remove(arquivo_completo)
                arquivos = {arq for arq in os.listdir(self.DIR_PEDIDOS) if arq.endswith(".csv")}

            if self.superus_iniciado:
                self.gerenciador_superus.fechar_processo()
                self.superus_iniciado = False

            if pdfs_cotacao:
                AutomacaoWhatsApp.processar_pdfs(pdfs_cotacao, "cotação")

            if pdfs_compras:
                AutomacaoWhatsApp.processar_pdfs(pdfs_compras, "compras")

            sleep(1)


if __name__ == "__main__":
    os.system("cls")
    gerenciador = GerenciadorDePedidos()
    gerenciador.processar()