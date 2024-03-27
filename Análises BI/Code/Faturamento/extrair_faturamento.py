import os
import subprocess
import pandas as pd
import pyautogui as pg
from time import sleep
from datetime import datetime, timedelta
from config import DIR_FATURAMENTO_SMJ, DIR_FATURAMENTO_STT, DIR_FATURAMENTO_VIX, DIR_FATURAMENTO_MCP, DIR_FATURAMENTO_TOTAL

class SuperusAutomacao:
    def iniciar(self):
        subprocess.Popen([r"C:\SUPERUS VIX\Superus.exe"])
        self.aguardar_abertura()

    def aguardar_abertura(self):
        while not pg.locateOnScreen(r"Imgs\superus.png", confidence=0.9):
            sleep(1)
        pg.write("848600"); sleep(3)
        pg.press("enter", presses=2, interval=1)
        while not pg.locateOnScreen(r"Imgs\superus_aberto.png", confidence=0.9):
            sleep(1)

    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

    def fechar(self):
        self.fechar_processo("SUPERUS.EXE")

class MenuSelecao:
    @staticmethod
    def obter_datas_formatadas():
        data_atual = datetime.now().date()
        dia_ontem = data_atual - timedelta(days=1)
        primeiro_dia_ano_passado = datetime(data_atual.year - 1, 1, 1).date()
        return primeiro_dia_ano_passado.strftime('%d%m%Y'), dia_ontem.strftime('%d%m%Y')

    @staticmethod
    def selecionar():
        primeiro_dia_ano_passado, ontem = MenuSelecao.obter_datas_formatadas()

        pg.click(246,32, duration=1)
        pg.click(265,81, duration=1)
        pg.click(522,161, duration=3)
        pg.click(326,109, duration=1)
        pg.press("up", presses=5)
        pg.press("enter"); sleep(2)

        pg.doubleClick(101,128, duration=1)
        pg.write(primeiro_dia_ano_passado)

        pg.doubleClick(249,130, duration=1)
        pg.write(ontem)

        # Lista de opções a serem selecionadas
        opcoes = [
            "TRESMANN - SMJ",
            "TRESMANN - STT",
            "TRESMANN - VIX",
            "TRESMANN - MCP"
        ]

        for opcao in opcoes:
            pg.click(362, 131, duration=1); sleep(30)
            pg.click(96,40, duration=1); sleep(5)

            pg.write(f"FATURAMENTO {opcao}")
            pg.press("enter")

            pg.click(326,109, duration=1)
            pg.press("down")
            pg.press("enter")

            sleep(1)

class ProcessarFaturamento:
    def __init__(self):
        self.arquivos = [DIR_FATURAMENTO_SMJ, DIR_FATURAMENTO_STT, DIR_FATURAMENTO_VIX, DIR_FATURAMENTO_MCP]
        self.substrings = {'SMJ': 'Faturamento SMJ', 'STT': 'Faturamento STT', 'VIX': 'Faturamento VIX', 'MCP': 'Faturamento MCP'}

    def merge_arquivos(self):
        df_final = pd.DataFrame()
        for arquivo in self.arquivos:
            df = pd.read_excel(arquivo, usecols=['Data', 'Total'])
            for substring, novo_nome in self.substrings.items():
                if substring in arquivo:
                    df.rename(columns={'Total': novo_nome}, inplace=True)
                    break
            else:
                df.rename(columns={'Total': 'Faturamento Total Geral'}, inplace=True)
            df_final = df if df_final.empty else pd.merge(df_final, df, on='Data', how='outer').fillna(0)
            os.remove(arquivo)
        df_final.sort_values(by='Data', inplace=True)
        df_final['Data'] = pd.to_datetime(df_final['Data']).dt.strftime('%d/%m/%Y')

        # Salvar o DataFrame usando ExcelWriter para ajustar as larguras das colunas
        writer = pd.ExcelWriter(DIR_FATURAMENTO_TOTAL, engine='xlsxwriter')
        df_final.to_excel(writer, sheet_name='FATURAMENTO TOTAL', index=False)
        self.ajustar_largura_colunas(writer, df_final)
        writer.close()

    @staticmethod
    def ajustar_largura_colunas(writer, df):
        worksheet = writer.sheets['FATURAMENTO TOTAL']
        for i, col in enumerate(df.columns):
            col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

class Extrairfaturamento:
    def extrair_faturamento(self):
        superus = SuperusAutomacao()
        superus.iniciar()
        menu = MenuSelecao()
        menu.selecionar()
        superus.fechar()
        processador = ProcessarFaturamento()
        processador.merge_arquivos()

if __name__ == "__main__":
    orquestrador = Extrairfaturamento()
    orquestrador.extrair_faturamento()