import csv
import os
import ctypes
import tkinter
import customtkinter
import pyautogui as pa
from time import sleep

# Definição das classes para gestão de arquivos e automação
class FileManager:
    def __init__(self, file_paths):
        self.file_paths = file_paths

    def write_to_file(self, file_path, data):
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file)
            for item in data:
                writer.writerow([item])

    def read_from_file(self, file_path):
        with open(file_path) as file:
            return file.readlines()

class InventoryAutomator:
    def __init__(self, caminho_csv):
        self.caminho_csv = caminho_csv

    @staticmethod
    def abrir_inventario():
        """Abre o inventário utilizando atalhos de teclado."""
        pa.hotkey('win', '3')
        sleep(1)
        pa.hotkey('ctrl', 'home')
        sleep(0.5)
        pa.press('alt')
        sleep(0.5)
        pa.press('e')
        sleep(0.5)
        pa.press('i')
        sleep(0.5)
        pa.press('f2')
        sleep(0.5)
        pa.press('tab')
        sleep(4)
        pa.press('down', presses=20)
        pa.press('up')
        pa.press('tab')

    @staticmethod
    def ajustar_caps_lock():
        """Ajusta o estado do Caps Lock para desligado."""
        if ctypes.windll.user32.GetKeyState(0x14):
            pa.press('capslock')

    def processar_codigo(self, codigo):
        """Digita e processa um código de produto no sistema."""
        pa.write(codigo)
        pa.press('enter')
        pa.press('enter')
        sleep(2)
        pa.press('left')
        pa.press('enter')
        sleep(1)

    def processar_inventario(self):
        """Processa todos os códigos de produto listados no arquivo CSV."""
        with open(self.caminho_csv, 'r') as arquivo:
            for linha in arquivo:
                codigo = linha.strip().split(';')[0]
                self.processar_codigo(codigo)

    def finalizar_inventario(self, nome_arquivo):
        """Finaliza o inventário, salvando com o nome especificado."""
        self.ajustar_caps_lock()
        pa.write(nome_arquivo.upper())
        sleep(1)
        pa.press('enter')
        sleep(1)
        pa.press('f9', presses=2, interval=1)

    def executar(self, nome_arquivo):
        """Executa o processo completo de automação de inventário."""
        self.abrir_inventario()
        self.processar_inventario()
        self.finalizar_inventario(nome_arquivo)

customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("INVENTÁRIOS")
        self.geometry("640x420")
        self.configure_gui()
        self.file_manager = FileManager({
            'SMJ': "F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\SMJ.csv",
            'STT': "F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\STT.csv",
            'VIX': "F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\VIX.csv",
            'Outros': "F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\Outros.csv"
        })

    def configure_gui(self):
        # Configuração da interface gráfica
        pass

    # Outros métodos da GUI

if __name__ == "__main__":
    app = App()
    app.mainloop()
