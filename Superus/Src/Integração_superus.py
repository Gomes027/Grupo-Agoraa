import os
import csv
import chardet
from time import sleep
from datetime import datetime
from datetime import datetime, timedelta
from Compras_e_cotações.Digitar_pedidos_v2_0 import digitar_pedido
from Transferências.Transferências_Simplificação import digitar_transferencia
from Notas_fiscais.Extrair_notas import extrair_notas_fiscais
from Historico_pedidos.Extrair_historico_de_pedidos import extrair_historico_pedidos
from Selects.Extrair_selects import executar_selects

def verificar_tipo_arquivo(arquivo_completo):
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

# Pasta de Pedidos
DIR_PEDIDOS = r"F:\COMPRAS\Automações.Compras\Fila de Pedidos"

# LOOP PRINCIPAL
if __name__ == "__main__":
    arquivos_existentes = set()
    ultima_execucao_extracao = datetime.now() - timedelta(minutes=5)

    while True:
        arquivos = set(os.listdir(DIR_PEDIDOS))
        novos_arquivos = arquivos - arquivos_existentes
        arquivos_transferencia, arquivos_cotacao, arquivos_compras = [], [], []

        for arquivo in novos_arquivos:
            if arquivo.endswith(".csv"):
                arquivo_completo = os.path.join(DIR_PEDIDOS, arquivo)
                tipo_pedido = verificar_tipo_arquivo(arquivo_completo)
                if tipo_pedido == "transferência":
                    arquivos_transferencia.append(arquivo)
                elif tipo_pedido == "cotação":
                    arquivos_cotacao.append(arquivo)
                elif tipo_pedido == "compras":
                    arquivos_compras.append(arquivo)

        # Processa os arquivos por tipo de pedido
        for arquivo in arquivos_transferencia + arquivos_cotacao + arquivos_compras:
            arquivo_completo = os.path.join(DIR_PEDIDOS, arquivo)
            tipo_pedido = verificar_tipo_arquivo(arquivo_completo)

            if tipo_pedido == "transferência":
                digitar_transferencia(arquivo, arquivo_completo)
            else:  # cotação e compras
                digitar_pedido(arquivo_completo, tipo_pedido, arquivo)

            os.remove(arquivo_completo)
        
        arquivos_existentes = arquivos
        23.300
        
        # Executa a extração de histórico de pedidos ou notas fiscais quando não houver arquivos novos
        if not novos_arquivos:
            agora = datetime.now()
            hora_atual = agora.strftime("%H:%M")

            if hora_atual == "20:00":
                extrair_historico_pedidos()
            elif hora_atual == "21:00":
                executar_selects()
                
            if 8 <= agora.hour < 18:
                # Chama extrair_notas_fiscais a cada 5 minutos
                agora = datetime.now()
                if agora - ultima_execucao_extracao >= timedelta(minutes=5):
                    extrair_notas_fiscais()
                    ultima_execucao_extracao = agora

        sleep(5)