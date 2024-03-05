import logging
from datetime import datetime
from Estoque_Excedente.gerar_estoques import atualizar_estoques
from Descontos.extrair_descontos import extrair_descontos
from Descontos.tabela_dinamica import tabela_dinamica
from Volume_dos_100.volume_dos_100 import volumes
from Encartes.extrair_encartes import extrair_encartes
from Encartes.relatorio_encartes import relatorio_encartes
from Lojas_de_Controle.extrair_lojas import extrair_lojas
from Lojas_de_Controle.relatorio_lojas import relatorio_lojas
from Vinhos.extrair_relatorios import extrair_relatorios
from Vinhos.relatorio_vinhos import gerar_relatorio_vinhos
from Faturamento.extrair_faturamento import extrair_faturamento
from Faturamento.relatorio_de_faturamento import gerar_relatorio_faturamento
from Email.enviar_email import enviar_email

# Configuração do logging
logging.basicConfig(filename=f'Logs\BI-{datetime.now().strftime("%d-%m-%Y")}.log', level=logging.INFO, filemode='w',
                    format='%(asctime)s:%(levelname)s:%(message)s', datefmt='%Y-%m-%d %H:%M:%S')

if __name__ == "__main__":
    try:
        logging.info("Iniciando atualização dos estoques")
        atualizar_estoques()
        logging.info("Atualização dos estoques concluída")

        logging.info("Iniciando extração de descontos")
        extrair_descontos()
        logging.info("Extração de descontos concluída")

        logging.info("Iniciando criação da tabela dinâmica")
        tabela_dinamica()
        logging.info("Criação da tabela dinâmica concluída")

        logging.info("Iniciando cálculo dos volumes dos 100")
        volumes()
        logging.info("Cálculo dos volumes dos 100 concluído")

        logging.info("Iniciando extração de encartes")
        extrair_encartes()
        logging.info("Extração de encartes concluída")

        logging.info("Iniciando geração do relatório de encartes")
        relatorio_encartes()
        logging.info("Geração do relatório de encartes concluída")

        logging.info("Iniciando extração de lojas")
        extrair_lojas()
        logging.info("Extração de lojas concluída")

        logging.info("Iniciando geração do relatório de lojas")
        relatorio_lojas()
        logging.info("Relatório de lojas gerado")

        logging.info("Iniciando extração de relatórios de vinhos")
        extrair_relatorios()
        logging.info("relatórios de vinhos extraidos")

        logging.info("Iniciando geração do relatório de vinhos")
        gerar_relatorio_vinhos()
        logging.info("relatório de vinhos gerado")

        logging.info("Iniciando extração dos relatórios de faturamentos")
        extrair_faturamento()
        logging.info("relatórios de faturamento extraidos")

        logging.info("Iniciando geração do relatório de faturamento")
        gerar_relatorio_faturamento()
        logging.info("relatório de faturamento gerado")

        logging.info("Iniciando envio de email")
        enviar_email()
        logging.info("Envio de email concluído")

    except Exception as e:
        logging.error("Erro no processo: %s", str(e))