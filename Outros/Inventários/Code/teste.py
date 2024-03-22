import os
import csv
import ctypes
import tkinter as tk
import pyautogui as pg
from time import sleep
from customtkinter import CTkFrame, CTkLabel, CTkButton, CTkOptionMenu, CTkTabview, CTkEntry, CTkRadioButton, CTkFont

class Inventarios:
    def __init__(self, caminho_outros):
        self.caminho_outros = caminho_outros

    def salvar_codigos(self, codigos):
        with open(self.caminho_outros, mode='w', newline='') as file:
            writer = csv.writer(file)
            for codigo in codigos:
                writer.writerow([codigo])

class Interface(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("INVENTÁRIOS")
        self.geometry(f"{640}x{420}")

        self.configurar_layout()
        self.criar_barra_lateral()
        self.criar_menu_seleção()
        self.entrada_codigos_e_salvar()
        self.create_radiobuttons()

    def configurar_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

    def criar_barra_lateral(self):
        self.barra_lateral = CTkFrame(self, width=140, corner_radius=0)
        self.barra_lateral.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.barra_lateral.grid_rowconfigure(4, weight=1)

        self.opções_label = CTkLabel(self.barra_lateral, text="Opções", font=CTkFont(family="Poppins", size=20, weight="bold"))
        self.opções_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        self.botão_iniciar = CTkButton(self.barra_lateral, text="Iniciar", fg_color="FireBrick", font=("Poppins", 12, "bold"), command=self.ação_botão_iniciar)
        self.botão_iniciar.grid(row=1, column=0, padx=20, pady=10)
        self.botão_reiniciar = CTkButton(self.barra_lateral, text="Reiniciar", fg_color="MidnightBlue", font=("Poppins", 12, "bold"), command=self.ação_botão_reiniciar)
        self.botão_reiniciar.grid(row=2, column=0, padx=20, pady=10)

        self.aparencia_label = CTkLabel(self.barra_lateral, text="Tema:", font=("Poppins", 12, "bold"), anchor="w")
        self.aparencia_label.grid(row=5, column=0, padx=20, pady=(10, 0))

        self.aparencia_menu = CTkOptionMenu(self.barra_lateral,  fg_color="MidnightBlue",  font=("Poppins", 12, "bold"), values=["Light", "Dark", "System"], command=self.ação_aparencia_menu)
        self.aparencia_menu.grid(row=6, column=0, padx=20, pady=(10, 10))
        
        self.escala_menu = CTkOptionMenu(self.barra_lateral,  fg_color="MidnightBlue", font=("Poppins", 12, "bold"), values=["80%", "90%", "100%", "110%", "120%"], command=self.ação_escala_menu)
        self.escala_menu.grid(row=8, column=0, padx=10, pady=(10, 20))

    def criar_menu_seleção(self):
        self.menu_seleção = CTkTabview(self, width=250)
        self.menu_seleção.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.menu_seleção.add("CONFIGURAÇÕES")
        self.menu_seleção.tab("CONFIGURAÇÕES").grid_columnconfigure(0, weight=1)

        self.numero_inventario = CTkOptionMenu(self.menu_seleção.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=[str(i) for i in range(1, 13)])
        self.numero_inventario.grid(row=1, column=0, padx=20, pady=(30, 10))

        self.quantidade_codigos = CTkOptionMenu(self.menu_seleção.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=[str(i) for i in range(1, 11)])
        self.quantidade_codigos.grid(row=3, column=0, padx=20, pady=(20, 10))

        self.fornecedor_selecionado = tk.StringVar()
        self.fornecedor = CTkOptionMenu(self.menu_seleção.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=("Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"),
                                                        variable=self.fornecedor_selecionado)

    def entrada_codigos_e_salvar(self):
        self.entrada_codigos = CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Codigos...")
        self.entrada_codigos.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.salvar_codigos = CTkButton(master=self, fg_color="green", text="Salvar", font=("Poppins", 12, "bold"), border_width=2, text_color=("gray10", "#DCE4EE"), command=self.save_text)
        self.salvar_codigos.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

    def create_radiobuttons(self):
        self.radiobutton_frame = CTkFrame(self)
        self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        # Adicione aqui a criação dos botões de rádio

    def save_text(self):
        # Adicione aqui o código para salvar o texto

    def optionmenu_3_var_changed(self, *args):
        # Adicione aqui o código para lidar com a mudança na opção do menu 3

    def ação_aparencia_menu(self, new_appearance_mode: str):
        # Adicione aqui o código para alterar o modo de aparência

    def ação_escala_menu(self, new_scaling: str):
        # Adicione aqui o código para alterar a escala dos widgets

    def ação_botão_iniciar(self):
        print("Inventário Iniciado")
        self.abrir_inventario()
        self.processar_inventario()
        self.finalizar_inventario()
        self.ajustar_interface()

    def abrir_inventario(self):
        pg.hotkey("win", "3"); sleep(1)
        pg.hotkey("ctrl", "home"); sleep(0.5)
        pg.press("alt"); sleep(0.5)
        pg.press("e"); sleep(0.5)
        pg.press("i"); sleep(0.5)
        pg.press("f2"); sleep(0.5)
        pg.press("tab"); sleep(4)
        pg.press("down", presses=20)
        pg.press("up")
        pg.press("tab")

    def processar_inventario(self):
        fornecedor_selecionado = self.optionmenu_3.get()
        if self.validar_inventario(fornecedor_selecionado):
            selected_value_numero_int = self.obter_numero_inventario()
            selected_value_final = self.obter_valor_final()
            caminho_csv = self.obter_caminho_csv(fornecedor_selecionado)
            self.executar_inventario(fornecedor_selecionado, caminho_csv, selected_value_final)

    def validar_inventario(self, fornecedor_selecionado):
        palavras_chave = ["Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"]
        if any(palavra in fornecedor_selecionado for palavra in palavras_chave):
            try:
                self.obter_numero_inventario()
                return True
            except ValueError:
                print("Escolha um Número Valido!")
                pg.press("f9")
                sleep(1)
                print('________________________________')
                return False
        else:
            return True

    def obter_numero_inventario(self):
        selected_value_numero = self.numero_inventario.get()
        selected_value_numero_int = int(selected_value_numero)
        print("Inventário Nº:", selected_value_numero_int)
        return selected_value_numero_int

    def obter_valor_final(self):
        selected_value_numero_invent = self.quantidade_codigos.get()
        if selected_value_numero_invent == "Quant. de Cod":
            return 6
        else:
            return int(selected_value_numero_invent)

    def obter_caminho_csv(self, fornecedor_selecionado):
        file_path = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv"
        if fornecedor_selecionado in ("Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"):
            caminho_csv = os.path.join(r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\Carnes e Paes", f"{fornecedor_selecionado} SMJ.csv")
        else:
            caminho_csv = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\Outros.csv"
            if os.path.exists(self.file_path) and os.path.getsize(self.file_path) == 0:
                if self.radio_var.get() == 0:
                    caminho_csv = self.caminho_SMJ
                elif self.radio_var.get() == 1:
                    caminho_csv = self.caminho_STT
                elif self.radio_var.get() == 2:
                    caminho_csv = self.caminho_VIX
        return caminho_csv

    def executar_inventario(self, fornecedor_selecionado, caminho_csv, selected_value_final):
        with open(caminho_csv) as arquivo:
            linhas = arquivo.readlines()
            sleep(2)

        contador = 0
        for linha in linhas:
            if self.checar_palavras_chave(fornecedor_selecionado):
                self.executar_codigo(linha)
            elif caminho_csv == r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\Outros.csv":
                self.executar_codigo(linha)
            else:
                if contador >= selected_value_final:
                    break
                self.executar_codigo(linha)
                contador += 1

    def checar_palavras_chave(self, fornecedor_selecionado):
        palavras_chave = ["Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"]
        return any(palavra in fornecedor_selecionado for palavra in palavras_chave)

    def executar_codigo(self, linha):
        codigo = linha.split(';')[0]
        pg.write(codigo)
        pg.press("enter")
        sleep(0.1)
        pg.hotkey("alt", "o")
        sleep(0.5)
        pg.click(615, 845, duration=0.5)
        sleep(0.1)
        pg.click(711, 371, duration=0.5)

    def finalizar_inventario(self):
        pg.press("esc")
        sleep(3)
        # Código para finalizar o inventário...

    def ajustar_interface(self):
        fornecedor_selecionado = self.fornecedor.get()
        if self.checar_palavras_chave(fornecedor_selecionado):
            self.resetar_interface()
        else:
            pass

    def resetar_interface(self):
        self.fornecedor.set("Carnes e Paes")
        self.numero_inventario.configure(state="normal")
        self.quantidade_codigos.configure(state="normal")
        self.salvar_codigos.configure(state="normal")
        self.entrada_codigos.configure(state="normal")
        self.radio_button_1.configure(state="normal")
        self.radio_button_2.configure(state="normal")

    def ação_botão_reiniciar(self):
        self.destroy()

        with open(self.file_path, mode='w', newline=''):
            pass

        app = App()
        app.mainloop()

class App:
    def __init__(self):
        self.caminho_SMJ = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\SMJ.csv"
        self.caminho_STT = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\STT.csv"
        self.caminho_VIX = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\VIX.csv"

        self.file_path = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Inventários\Outros.csv"
        self.inventarios = Inventarios(self.file_path)
        self.interface = Interface(self.inventarios)

    def run(self):
        self.interface.mainloop()

if __name__ == "__main__":
    app = App()
    app.run()