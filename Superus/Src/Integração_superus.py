import os
import csv
import chardet
import pandas as pd
from time import sleep
import pyautogui as pg
from datetime import datetime, timedelta
from Selects.Extrair_selects import executar_selects
from Notas_fiscais.Extrair_notas import extrair_notas_fiscais
from Compras_e_cotações.Digitar_pedidos_v2_0 import digitar_pedido
from Compras_e_cotações.Enviar_Pedidos_WhatsApp import procesar_pdfs
from Transferências.Transferências_Simplificação import digitar_transferencia
from Historico_pedidos.Extrair_historico_de_pedidos import extrair_historico_pedidos
from Validades.src.Calculadora_de_validades import processar_arquivo

# Constantes
DIR_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos"
DIR_PDF_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Compras\PDFs"
DIR_PDF_COTACAO = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Cotações\PDFs"

def verificar_tipo_arquivo(arquivo_completo):
    """
    Determina o tipo de um arquivo com base no seu conteúdo.
    """
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

def pontuacao_tipo_arquivo(tipo_pedido):
    """                                                     
    Atribui uma pontuação a cada tipo de pedido para fins de ordenação.
    """
    return {"transferência": 1, "cotação": 2, "compras": 3}.get(tipo_pedido, 4)

primeira_vez = True

# Loop principal
if __name__ == "__main__":
    ultima_execucao_extracao = datetime.now() - timedelta(minutes=5)

    while True:
        validades = {arq for arq in os.listdir(DIR_PEDIDOS) if arq.endswith(".xlsx")}
        arquivos = {arq for arq in os.listdir(DIR_PEDIDOS) if arq.endswith(".csv")}
        pdfs_cotacao = {arq for arq in os.listdir(DIR_PDF_COTACAO) if arq.lower().endswith(".pdf")}
        pdfs_compras = {arq for arq in os.listdir(DIR_PDF_PEDIDOS) if arq.lower().endswith(".pdf")}

        if validades:
            for validade in validades:
                caminho_validade = os.path.join(DIR_PEDIDOS, validade)
                processar_arquivo(caminho_validade)
                os.remove(caminho_validade)

        elif arquivos:
            arquivos_ordenados = sorted(
                arquivos, 
                key=lambda arq: (pontuacao_tipo_arquivo(verificar_tipo_arquivo(os.path.join(DIR_PEDIDOS, arq))), arq)
            )

            for arquivo in arquivos_ordenados:
                arquivo_completo = os.path.join(DIR_PEDIDOS, arquivo)
                tipo_pedido = verificar_tipo_arquivo(arquivo_completo)
                
                if tipo_pedido == "transferência":
                    digitar_transferencia(arquivo, arquivo_completo)
                elif tipo_pedido in ["cotação", "compras"]:
                    digitar_pedido(arquivo_completo, tipo_pedido, arquivo)

                os.remove(arquivo_completo)
                break

        elif pdfs_cotacao:
            # Processamento dos PDFs de cotação
            procesar_pdfs(pdfs_cotacao, "cotação")

        elif pdfs_compras:
            # Processamento dos PDFs de compras
            procesar_pdfs(pdfs_compras, "compras")       
                
        sleep(1)