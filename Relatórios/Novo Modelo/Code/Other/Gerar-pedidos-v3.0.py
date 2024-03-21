import os
import re
import csv
import sys
import shutil
import keyboard
import pyperclip
import subprocess
from time import sleep
import pyautogui as pg
import pygetwindow as gw
from PIL import ImageGrab
from config import *

class ClipboardManager:
    @staticmethod
    def aguardar_conteudo_area_transferencia(atalho, tempo_espera=60, tentativas=5):
        while tentativas > 0:
            pyperclip.copy("")  # Limpa a área de transferência
            conteudo_inicial = pyperclip.paste()  # Guarda o conteúdo inicial (vazio)

            ClipboardManager.executar_atalho(atalho)

            tempo_decorrido = 0
            while tempo_decorrido < tempo_espera:
                sleep(1)  # Aguarda um segundo
                tempo_decorrido += 1
                conteudo_atual = pyperclip.paste()
                if conteudo_atual != conteudo_inicial:
                    return conteudo_atual  # Retorna o novo conteúdo

            print("Tempo de espera excedido. Tentativa:", tentativas)
            tentativas -= 1  # Decrementa o contador de tentativas

        print("Nenhum conteúdo novo na área de transferência após as tentativas.")
        return None

    @staticmethod
    def executar_atalho(atalho):
        atalhos = {
            "Nome do Pedido": lambda: pg.hotkey("ctrl", "shift", "l"),
            "Dados do Pedido": lambda: pg.hotkey("ctrl", "shift", "r"),
            "CSV do Pedido": lambda: pg.hotkey("ctrl", "shift", "t"),
            "Imagem do Pedido": lambda: pg.hotkey("ctrl", "shift", "s"),
            "Pedidos Gerados": lambda: pg.hotkey("ctrl", "shift", "h"),
            "Nome do Fornecedor": lambda: pg.hotkey("ctrl", "shift", "y"),
        }
        if atalho in atalhos:
            atalhos[atalho]()

class FileManager:
    @staticmethod
    def limpar_dados(dados):
        """Limpa o nome do arquivo removendo caracteres especiais."""
        return re.sub(r'[<>:"/\\|?*]', '', dados).strip()

    @staticmethod
    def salvar_em_csv(caminho_da_pasta, nome_arquivo, dados):
        # Limpa o nome do arquivo usando a função limpar_dados que já definimos
        nome_arquivo_limpo = FileManager.limpar_dados(nome_arquivo)
        caminho_completo = os.path.join(caminho_da_pasta, f"{nome_arquivo_limpo}.csv")
        os.makedirs(os.path.dirname(caminho_completo), exist_ok=True)

        linhas = [dados[i:i + 3] for i in range(0, len(dados), 3)]

        # Escrever as linhas do arquivo CSV
        with open(caminho_completo, 'w', newline='', encoding='utf-8') as csvfile:
            csv_writer = csv.writer(csvfile, delimiter=';')
            for linha in linhas:
                csv_writer.writerow(linha)
        
        # Reabrir o arquivo para filtrar linhas válidas
        linhas_validas = []
        with open(caminho_completo, mode='r', newline='', encoding='utf-8') as file:
            reader = csv.reader(file, delimiter=';')
            for linha in reader:
                if len(linha) > 1 and linha[1] != '0' and linha[1] != '':
                    linhas_validas.append(linha)

        # Reescrever o arquivo com apenas as linhas válidas
        with open(caminho_completo, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file, delimiter=';')
            writer.writerows(linhas_validas)

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

    @staticmethod
    def processar_imagens(pasta):
        WindowManager.abrir_janela("WhatsApp")
        WindowManager.procurar_botao_e_clicar(r"grupo_jaidson.png", semelhança=0.9)
        WindowManager.procurar_botao_e_clicar(r"mais.png")
        WindowManager.procurar_botao_e_clicar(r"fotos.png")
        WindowManager.procurar_botao_e_clicar(r"explorer.png", clicar=False); sleep(5)

        keyboard.send("ctrl+l"); sleep(1)
        keyboard.write(pasta); pg.press("enter")
        pg.press("tab", presses=4); sleep(1)
        pg.hotkey("ctrl", "a"); sleep(1)
        pg.press("enter")

        WindowManager.procurar_botao_e_clicar(r"enviar.png", clicar=False); sleep(10)
        pg.click(572, 597, duration=1); sleep(3)

        arquivos = [arq for arq in os.listdir(pasta) if os.path.splitext(arq)[1].lower() in ['.png', '.jpg', '.jpeg']]
        for imagem in arquivos:
            nome_sem_extensao = os.path.splitext(imagem)[0]
            keyboard.write(nome_sem_extensao)
            sleep(3)
            pg.press("right")
            sleep(2)
            print(f"Processando imagem: {nome_sem_extensao}")

        sleep(30); pg.press("enter")
        WindowManager.procurar_botao_e_clicar(r"digitar.png", semelhança=0.7); sleep(3)
        pg.typewrite("*PEDIDOS FINALIZADOS*"); sleep(3); pg.press("enter")

        for imagem in arquivos:
            os.remove(os.path.join(pasta, imagem))

class WindowManager:
    @staticmethod
    def procurar_botao_e_clicar(imagem, clicar=True, semelhança=0.8):
        while True:
            try:
                localizacao = pg.locateOnScreen(os.path.join("Imgs", imagem),confidence=semelhança)
                if localizacao:
                    if clicar:
                        pg.click(pg.center(localizacao), duration=0.5)
                    break
            except pg.ImageNotFoundException:
                print(f"Imagem {imagem} não encontrada. Tentando novamente...")
            except Exception as e:
                print(f"Erro ao procurar a imagem {imagem}: {e}")

            sleep(3)

    @staticmethod
    def excel_modo_leitura(caminho_planilha):
        caminho_excel = r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE"

        comando_planilha = f'"{caminho_excel}" /r "{caminho_planilha}"'
        subprocess.Popen(comando_planilha, shell=True)
        sleep(10)

    @staticmethod
    def fechar_processo(nome_do_processo):
        try:
            subprocess.run(['taskkill', '/f', '/im', nome_do_processo], check=True)
        except subprocess.CalledProcessError:
            pass

    @staticmethod            
    def abrir_janela(nome_da_janela):
        try:
            janelas = gw.getWindowsWithTitle(nome_da_janela)
            if janelas:
                janela = janelas[0]
                if janela.isMinimized:
                    janela.restore()
                janela.activate()
                sleep(1)
            else:
                print(f"Janela {nome_da_janela} não encontrada.")
                sys.exit()
        except Exception as e:
            print(f"Erro ao ativar a janela: {e}")
            sys.exit()

class ExcelDataManager:
    def __init__(self, relatorio, clipboard_manager):
        self.relatorio = relatorio
        self.clipboard_manager = clipboard_manager
        self.image_processor = ImageProcessor()

    def preencher_dados_excel(self):
        with open(self.relatorio, newline='') as arquivo:
            leitor = csv.reader(arquivo, delimiter=';')
            linhas_ordenadas = sorted(leitor, key=lambda row: row[2])  # Ordena pelo índice 2 (terceira coluna)

        loja_macro_anterior = None

        for linha in linhas_ordenadas:
            fornecedor, loja, loja_macro = linha[0], linha[1], linha[2].strip()
            self._selecionar_fornecedor(fornecedor)
            self._usar_macro_para_lojas(loja_macro, loja_macro_anterior)
            loja_macro_anterior = loja_macro

            pg.hotkey('ctrl', 'shift', 'j')
            sleep(0.5)          
            pg.press('tab', presses=2)
            pg.press('enter')  
            sleep(7)

            self.processar_dados_pedido(fornecedor, loja)
            
        self.clipboard_manager.executar_atalho("Pedidos Gerados")

    def _selecionar_fornecedor(self, fornecedor):
        self.clipboard_manager.executar_atalho("Nome do Fornecedor"); sleep(3)
        pg.press("space"); pg.press("backspace")
        keyboard.write(fornecedor)
        sleep(0.5)
        pg.press('enter')
        sleep(7)

    def _usar_macro_para_lojas(self, loja_macro, loja_macro_anterior):
        if loja_macro != loja_macro_anterior:
            sleep(20)
            macro_comandos = {
                'M': "ctrl+shift+m", 'N': "ctrl+shift+n", 'B': "ctrl+shift+b", 'I': "ctrl+shift+i",
                'X': "ctrl+shift+x", 'E': "ctrl+shift+e", 'W': "ctrl+shift+w", 'V': "ctrl+shift+v",
                'Z': "ctrl+shift+z", 'G': "ctrl+shift+g", 'D': "ctrl+shift+d"
            }
            comando = macro_comandos.get(loja_macro.strip())
            if comando:
                pg.hotkey(*comando.split('+'))
                sleep(50)
            else:
                print("Lojas erradas, a digitação sera encerrada!")
                sys.exit()

    def processar_dados_pedido(self, fornecedor, loja):
        dados = self.clipboard_manager.aguardar_conteudo_area_transferencia("Dados do Pedido")
        elementos = dados.split()

        if len(elementos) >= 3:
            comprador, valor_total, valor_minimo = elementos[:3]
            print("Fornecedor:", fornecedor)
            print("Loja:", loja)
            print("Comprador:", comprador)
            print("Valor Total:", valor_total)
            print("Valor Mínimo:", valor_minimo)

            if valor_total == "#REF!":
                pg.hotkey('ctrl', 'shift', 'k')
                nome_img = f"{fornecedor} {loja} - NÃO DEU PEDIDO"
            elif float(valor_total) < float(valor_minimo):
                pg.hotkey('ctrl', 'shift', 'k')
                nome_img = f"{fornecedor} {loja} - NÃO DEU MÍNIMO R${valor_total}"
            else:
                nome_img = f"{fornecedor} {loja} R${valor_total}"

            if valor_total != "#REF!":
                nome_arquivo = self.clipboard_manager.aguardar_conteudo_area_transferencia("Nome do Pedido")
                dados_csv = self.clipboard_manager.aguardar_conteudo_area_transferencia("CSV do Pedido")
                if dados_csv:
                    FileManager.salvar_em_csv(DIR_CSV, nome_arquivo, dados_csv.split())

            caminho_completo = os.path.join(DIR_JAIDSON if comprador == "JAIDSON" else DIR_OUTROS, f"{nome_img}.png")
            self.clipboard_manager.aguardar_conteudo_area_transferencia("Imagem do Pedido")
            ImageProcessor.salvar_imagem(self, caminho_completo, ImageGrab.grabclipboard())
        else:
            print("Dados insuficientes") 

class AutomationManager:
    def __init__(self):
        self.clipboard_manager = ClipboardManager()
        self.file_manager = FileManager()
        self.image_processor = ImageProcessor()
        self.window_manager = WindowManager()
        
    def processar_arquivo(self, arquivo_completo):
        # Certifica-se de que a janela correta está aberta antes de processar
        keyboard.wait("insert")

        # Processar os dados no Excel
        excel_data_manager = ExcelDataManager(arquivo_completo, self.clipboard_manager)
        excel_data_manager.preencher_dados_excel()

        # Mover o arquivo para a pasta de arquivos processados
        self._mover_para_relatorios_antigos(arquivo_completo)

    def _mover_para_relatorios_antigos(self, arquivo_completo):
        try:
            relatorios_antigos = os.path.join(DIR_RELATORIOS, "Antigos")
            shutil.move(arquivo_completo, relatorios_antigos)
            print(f"Arquivo {arquivo_completo} movido para {self.relatorio_antigos}.")
        except Exception as e:
            print(f"Erro ao mover o arquivo {arquivo_completo}: {e}")

    def iniciar(self):
        while True:
            arquivos_existentes = set()
            while True:
                arquivos = set(os.listdir(DIR_RELATORIOS))
                novos_arquivos = arquivos - arquivos_existentes
                novos_arquivos = [arq for arq in sorted(novos_arquivos) if arq.endswith(".csv")]

                for novo_arquivo in novos_arquivos:
                    self.window_manager.excel_modo_leitura(DIR_CONTROLE_DE_COMPRA)
                    self.window_manager.excel_modo_leitura(DIR_NOVO_MODELO_COMPRAS)

                    arquivo_completo = os.path.join(DIR_RELATORIOS, novo_arquivo)
                    self.processar_arquivo(arquivo_completo)
                    self.window_manager.fechar_processo("EXCEL.EXE")
                    self.image_processor.processar_imagens(r"F:\COMPRAS\Automações.Compras\Fila de Pedidos\Arquivos\Novo Modelo\Imgs\Jaidson")

                arquivos_existentes.update(novos_arquivos)
                sleep(1)  # Evitar sobrecarga de CPU e permitir novos arquivos serem adicionados

if __name__ == "__main__":
    automation_manager = AutomationManager()
    automation_manager.iniciar()