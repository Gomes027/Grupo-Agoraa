import os
import csv
import sys
import locale
import shutil
import getpass
import keyboard
import pyperclip
import subprocess
import pandas as pd
from time import sleep
import pyautogui as pg
import pygetwindow as gw
from PIL import ImageGrab
from Novo_modelo.config import *
from Novo_modelo.Enviar_pedidos import EnviarPedidos

class ClipboardManager:
    @staticmethod
    def aguardar_conteudo_area_transferencia(atalho, tempo_espera=60, tentativas=3):
        while tentativas > 0:
            pyperclip.copy("")
            conteudo_inicial = pyperclip.paste()

            ClipboardManager.executar_atalho(atalho)

            tempo_decorrido = 0
            while tempo_decorrido < tempo_espera:
                sleep(1)
                tempo_decorrido += 1
                conteudo_atual = pyperclip.paste()
                if conteudo_atual != conteudo_inicial:
                    return conteudo_atual 

            print("Tempo de espera excedido. Tentativa:", tentativas)
            tentativas -= 1

        print("Nenhum conteúdo novo na área de transferência após as tentativas."); sys.exit()

    @staticmethod
    def executar_atalho(atalho):
        atalhos = {
            "Imagem do Pedido": lambda: pg.hotkey("ctrl", "shift", "s"),
            "Pedidos Gerados": lambda: pg.hotkey("ctrl", "shift", "h"),
            "Nome do Fornecedor": lambda: pg.hotkey("ctrl", "shift", "y"),
        }
        if atalho in atalhos:
            atalhos[atalho]()

class PedidoProcessor:
    def __init__(self, novo_arquivo):
        self.novo_arquivo = novo_arquivo
        self.df_csv = pd.read_csv(self.novo_arquivo, header=None, sep=";")
        self.df_resultado = pd.DataFrame(columns=['FORNECEDOR', 'LOJA', 'VALOR TOTAL', 'RESULTADO', 'COMPRADOR'])
        self.df_novo_modelo = pd.read_excel(DIR_NOVO_MODELO_COMPRAS, sheet_name='relatorio_tresmann', header=2)
        self.df_controle_compra = pd.read_excel(DIR_CONTROLE_DE_COMPRA, sheet_name='BASE', header=1)
        self.df_comprador = pd.read_excel(DIR_CONTROLE_DE_COMPRA, sheet_name='COMPRADOR')
        self.lojas_dict = {
            'VIX (+SMJ+STT+MCP)': ['TRESMANN - VIX', 'TRESMANN - SMJ', 'TRESMANN - STT', 'MERCAPP'],
            'SMJ (+VIX+STT+MCP)': ['TRESMANN - SMJ', 'TRESMANN - VIX', 'TRESMANN - STT', 'MERCAPP'],
            'STT (+VIX+SMJ+MCP)': ['TRESMANN - STT', 'TRESMANN - VIX', 'TRESMANN - SMJ', 'MERCAPP'],
            'MCP (+VIX+SMJ+STT)': ['MERCAPP', 'TRESMANN - VIX', 'TRESMANN - SMJ', 'TRESMANN - STT'],
            'SMJ (+VIX+STT)': ['TRESMANN - SMJ', 'TRESMANN - VIX', 'TRESMANN - STT'],
            'VIX (+SMJ+STT)': ['TRESMANN - VIX', 'TRESMANN - SMJ', 'TRESMANN - STT'],
            'VIX (+SMJ+MCP)': ['TRESMANN - VIX', 'TRESMANN - SMJ', 'MERCAPP'],
            'SMJ (+VIX+MCP)': ['TRESMANN - SMJ', 'TRESMANN - VIX', 'MERCAPP'],
            'STT (+VIX+MCP)': ['TRESMANN - STT', 'TRESMANN - VIX', 'MERCAPP'],
            'MCP (+VIX+STT)': ['MERCAPP', 'TRESMANN - VIX', 'TRESMANN - STT'],
            'SMJ (+STT)': ['TRESMANN - SMJ', 'TRESMANN - STT'],
            'SMJ (+VIX)': ['TRESMANN - SMJ', 'TRESMANN - VIX'],
            'VIX (+MCP)': ['TRESMANN - VIX', 'MERCAPP'],
            'SMJ (+MCP)': ['TRESMANN - SMJ', 'MERCAPP'],
            'STT (+MCP)': ['TRESMANN - STT', 'MERCAPP'],
            'SMJ': 'TRESMANN - SMJ',
            'STT': 'TRESMANN - STT',
            'VIX': 'TRESMANN - VIX',
            'MCP': 'MERCAPP'
        }

    def obter_nome_loja(self, loja_abreviada):
        lojas = self.lojas_dict.get(loja_abreviada, '')
        if isinstance(lojas, list):
            return lojas
        else:
            return [lojas]

    def processar_pedidos(self):
        for index, row in self.df_csv.iterrows():
            fornecedor = row[0]
            loja_abreviada = row[1]
            nome_loja = self.obter_nome_loja(loja_abreviada)
            filtro = (self.df_novo_modelo['NOME_FANTASIA'] == fornecedor) & (self.df_novo_modelo['LOJA'].apply(lambda x: any(l in x for l in nome_loja))) & (self.df_novo_modelo['FORA DO MIX'] == "NAO")
            total_valor_novo = self.df_novo_modelo.loc[filtro, 'VALOR NOVO'].sum()
            self.df_resultado = pd.concat([self.df_resultado, pd.DataFrame({'FORNECEDOR': [fornecedor], 'LOJA': [loja_abreviada], 'VALOR TOTAL': [total_valor_novo]})], ignore_index=True)

    def verificar_resultados(self):
        resultados = []
        comprador = []
        for index, row in self.df_resultado.iterrows():
            fornecedor = row['FORNECEDOR']
            valor_total = row['VALOR TOTAL']
            filtro_controle = (self.df_controle_compra['FORNECEDOR'] == fornecedor)
            pd_minimo = self.df_controle_compra.loc[filtro_controle, 'PD. MÍNIMO'].sum()
            comprador_por_fornecedor = self.df_comprador.loc[filtro_controle, 'COMPRADOR'].iloc[0]
            if valor_total == 0:
                resultado = 'NÃO DEU PEDIDO'
            elif valor_total > pd_minimo:
                resultado = 'DEU PEDIDO'
            else:
                resultado = 'NÃO DEU MÍNIMO'
            resultados.append(resultado)
            comprador.append(comprador_por_fornecedor)
        self.df_resultado['RESULTADO'] = resultados
        self.df_resultado['COMPRADOR'] = comprador

    def salvar_resultados(self):
        for index, row in self.df_resultado.iterrows():
            fornecedor = row['FORNECEDOR']
            loja = row['LOJA']
            primeira_loja = loja.split(' ')[0]
            comprador = row['COMPRADOR']
            nome_loja = self.obter_nome_loja(loja)
            filtro_controle = (self.df_novo_modelo['NOME_FANTASIA'] == fornecedor) & (self.df_novo_modelo['LOJA'].apply(lambda x: any(l in x for l in nome_loja)))
            df_resultado_fornecedor_loja = self.df_novo_modelo.loc[filtro_controle, ['CODIGO', 'NOVO CHECK COMPRA', 'QTDE NA EMBALAGEM']]
            df_resultado_fornecedor_loja['NOVO CHECK COMPRA'] = df_resultado_fornecedor_loja['NOVO CHECK COMPRA'].astype(int)
            df_resultado_fornecedor_loja = df_resultado_fornecedor_loja[df_resultado_fornecedor_loja['NOVO CHECK COMPRA'] != 0]
            nome_arquivo = f'{fornecedor} {primeira_loja} {comprador}.csv'
            caminho_arquivo_novo = os.path.join(DIR_CSV, nome_arquivo)
            if not df_resultado_fornecedor_loja.empty:
                df_resultado_fornecedor_loja.to_csv(caminho_arquivo_novo, index=False, sep=";", header=None)
        self.df_resultado.to_csv(self.novo_arquivo, index=False, sep=";")

class ImageProcessor:
    @staticmethod
    def salvar_imagem(self, caminho_completo, imagem):
        try:
            imagem = ImageGrab.grabclipboard()
            if imagem:
                imagem.save(caminho_completo, 'PNG')
                print(f"Imagem salva em: {caminho_completo}")
            else:
                print("Não há imagem na área de transferência.")
        except Exception as e:
            print(f"Ocorreu um erro: {e}")
        print("\n")

class WindowManager:
    @staticmethod
    def excel_modo_leitura(caminho_planilha):
        comando_planilha = f'"{DIR_EXCEL}" /r "{caminho_planilha}"'
        subprocess.Popen(comando_planilha, shell=True)
        sleep(10)

    @staticmethod
    def abrir_janela(nome_da_janela, intervalo_verificacao=10):
        """ Seleciona a janela especificada e a ativa. """
        while True:
            try:
                janelas = gw.getWindowsWithTitle(nome_da_janela)
                if janelas:
                    janela = janelas[0]
                    if janela.isMinimized:
                        janela.restore()
                    janela.activate()
                    sleep(intervalo_verificacao)
                    return  # Retorna se a janela for aberta com sucesso
                else:
                    print(f"Janela {nome_da_janela} não encontrada. Tentando novamente em {intervalo_verificacao} segundos.")
                    sleep(intervalo_verificacao)
            except Exception as e:
                print(f"Erro ao ativar a janela: {e}")
                sleep(intervalo_verificacao)

    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

class ExcelDataManager:
    def __init__(self, relatorio, clipboard_manager):
        self.relatorio = relatorio
        self.usuario_atual = getpass.getuser()
        self.clipboard_manager = clipboard_manager
        self.image_processor = ImageProcessor()

    def preencher_dados_excel(self):
        with open(self.relatorio, newline='', encoding='utf-8') as arquivo:
            leitor = csv.DictReader(arquivo, delimiter=';')
            linhas_ordenadas = sorted(leitor, key=lambda row: row["LOJA"])

        loja_anterior = None

        for linha in linhas_ordenadas:
            fornecedor, loja, valor_total, resultado, comprador = linha["FORNECEDOR"], linha["LOJA"], linha["VALOR TOTAL"], linha["RESULTADO"], linha["COMPRADOR"]

            # Verificar se o valor total é zero
            if float(valor_total) == 0:
                valor_total = ""
            else:
                # Converter o valor total para float e formatar como moeda local
                valor_total = f"R${locale.format_string('%d', float(valor_total), grouping=True)}"
            
            print(f"Fornecedor: {fornecedor}, Loja: {loja}, Valor Total: {valor_total}, Resultado: {resultado}, Comprador: {comprador}")

            self._selecionar_fornecedor(fornecedor)
            self._usar_macro_para_lojas(loja, loja_anterior)
            loja_anterior = loja

            self.processar_dados_pedido(fornecedor, loja, valor_total, resultado, comprador)
            
        self.clipboard_manager.executar_atalho("Pedidos Gerados"); sleep(10)

    def _selecionar_fornecedor(self, fornecedor):
        self.clipboard_manager.executar_atalho("Nome do Fornecedor"); sleep(3)
        pg.press("space"); pg.press("backspace")
        keyboard.write(fornecedor)
        pg.press('enter')

    def _usar_macro_para_lojas(self, loja, loja_anterior):
        if loja != loja_anterior:
            macro_comandos = {
                'SMJ': "ctrl+shift+m",
                'STT': "ctrl+shift+n",
                'VIX': "ctrl+shift+b",
                'SMJ (+MCP)': "ctrl+shift+d",
                'SMJ (+STT)': "ctrl+shift+i",
                'SMJ (+VIX)': "ctrl+shift+x",
                'SMJ (+STT+MCP)': "ctrl+shift+e",
                'SMJ (+VIX+MCP)': "ctrl+shift+w",
                'SMJ (+STT+VIX)': "ctrl+shift+v",
                'SMJ (+VIX+STT)': "ctrl+shift+v",
                'VIX (+SMJ+STT)': "ctrl+shift+v",
                'SMJ (+VIX+STT+MCP)': "ctrl+shift+z",
                'VIX (+SMJ+STT+MCP)': "ctrl+shift+z"
            }
            comando = macro_comandos.get(loja.strip())
            if comando:
                pg.hotkey(*comando.split('+'))
            else:
                print("Lojas erradas, a digitação sera encerrada!")
                sys.exit()

    def processar_dados_pedido(self, fornecedor, loja, valor_total, resultado, comprador):
        if resultado in ("NÃO DEU PEDIDO", "NÃO DEU MÍNIMO"):
            nome_img = f"{fornecedor} {loja} - {resultado} {valor_total}"
            pg.hotkey('ctrl', 'shift', 'k')
        else:
            nome_img = f"{fornecedor} {loja} {valor_total}"
            pg.hotkey('ctrl', 'shift', 'j')

        caminho_imgs = os.path.join(DIR_IMGS, self.usuario_atual, comprador, f"{nome_img}.png")
        if not os.path.exists(os.path.dirname(caminho_imgs)):
            try:
                os.makedirs(os.path.dirname(caminho_imgs))
            except Exception as e:
                print(f"Erro ao criar diretório: {e}")
                sys.exit()

        self.clipboard_manager.aguardar_conteudo_area_transferencia("Imagem do Pedido")
        ImageProcessor.salvar_imagem(self, caminho_imgs, ImageGrab.grabclipboard())

class AutomationManager:
    def __init__(self):
        self.usuario_atual = getpass.getuser()
        self.clipboard_manager = ClipboardManager()
        self.image_processor = ImageProcessor()
        self.window_manager = WindowManager()
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
        
    def processar_arquivo(self, arquivo_completo):
        # Processar os dados no Excel
        excel_data_manager = ExcelDataManager(arquivo_completo, self.clipboard_manager)
        excel_data_manager.preencher_dados_excel()

    def mover_para_relatorios_antigos(self, arquivo_completo):
        try:
            relatorios_antigos = os.path.join(DIR_RELATORIOS)
            shutil.move(arquivo_completo, relatorios_antigos)
            print(f"Arquivo {arquivo_completo} movido para {self.relatorio_antigos}.")
        except Exception as e:
            print(f"Erro ao mover o arquivo {arquivo_completo}: {e}")

    def iniciar(self, arquivo_completo):
        self.pedido_processor = PedidoProcessor(arquivo_completo)
        self.mover_para_relatorios_antigos(arquivo_completo)
        
        self.pedido_processor.processar_pedidos()
        self.pedido_processor.verificar_resultados()
        self.pedido_processor.salvar_resultados()

        self.window_manager.excel_modo_leitura(DIR_CONTROLE_DE_COMPRA)
        self.window_manager.excel_modo_leitura(DIR_NOVO_MODELO_COMPRAS)
        self.window_manager.abrir_janela("NOVO MODELO COMPRAS v10.20")

        self.processar_arquivo(arquivo_completo)
        self.window_manager.fechar_processo("EXCEL.EXE"); sleep(5)

        enviar_pedidos = EnviarPedidos(os.path.join(DIR_IMGS, self.usuario_atual))
        enviar_pedidos.processar_imagens("JAIDSON")

if __name__ == "__main__":
    automation_manager = AutomationManager()
    automation_manager.iniciar()