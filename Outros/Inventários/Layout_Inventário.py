import csv
import os
import ctypes
import tkinter
import customtkinter
import pyautogui as pa
from time import sleep

caminho_SMJ = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\SMJ.csv"
caminho_STT = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\STT.csv"
caminho_VIX = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\VIX.csv"

customtkinter.set_appearance_mode("Dark")  # Modos: "System" (padrão), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Temas: "blue" (padrão), "green", "dark-blue"

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title("INVENTÁRIOS")
        self.geometry(f"{640}x{420}")

        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Opções", font=customtkinter.CTkFont(family="Poppins", size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text="Iniciar", fg_color="FireBrick", font=("Poppins", 12, "bold"), command=self.sidebar_button_event_1)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text="Reiniciar", fg_color="MidnightBlue", font=("Poppins", 12, "bold"), command=self.sidebar_button_event_2)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Tema:", font=("Poppins", 12, "bold"), anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,  fg_color="MidnightBlue",  font=("Poppins", 12, "bold"), values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame,  fg_color="MidnightBlue", font=("Poppins", 12, "bold"), values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=10, pady=(10, 20))

        self.tabview = customtkinter.CTkTabview(self, width=250)
        self.tabview.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.add("CONFIGURAÇÕES")
        self.tabview.tab("CONFIGURAÇÕES").grid_columnconfigure(0, weight=1)

        self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=[str(i) for i in range(1, 13)])
        self.optionmenu_1.grid(row=1, column=0, padx=20, pady=(30, 10))

        self.optionmenu_2 = customtkinter.CTkOptionMenu(self.tabview.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=[str(i) for i in range(1, 11)])
        self.optionmenu_2.grid(row=3, column=0, padx=20, pady=(20, 10))

        self.selected_fornecedor = tkinter.StringVar()
        self.optionmenu_3 = customtkinter.CTkOptionMenu(self.tabview.tab("CONFIGURAÇÕES"), font=("Poppins", 12, "bold"), dynamic_resizing=False,
                                                        values=("Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"),
                                                        variable=self.selected_fornecedor)
        self.optionmenu_3.grid(row=5, column=0, padx=20, pady=(20, 10))
        
        self.selected_fornecedor.trace("w", self.optionmenu_3_var_changed)

        self.entry = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Codigos...")
        self.entry.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 20), sticky="nsew")

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="green", text="Salvar", font=("Poppins", 12, "bold"), border_width=2, text_color=("gray10", "#DCE4EE"), command=self.save_text)
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 20), sticky="nsew")

        self.radiobutton_frame = customtkinter.CTkFrame(self)
        self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.radio_var = tkinter.IntVar(value=0)
        self.label_radio_group = customtkinter.CTkLabel(master=self.radiobutton_frame, font=("Poppins", 14, "bold"), text="Lojas")
        self.label_radio_group.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, font=("Poppins", 12, "bold"), text= "SMJ", variable=self.radio_var, value=0)
        self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")
        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, font=("Poppins", 12, "bold"), text= "STT", variable=self.radio_var, value=1)
        self.radio_button_2.grid(row=3, column=2, pady=10, padx=20, sticky="n")
        self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, font=("Poppins", 12, "bold"), text= "VIX", variable=self.radio_var, value=2)
        self.radio_button_3.grid(row=5, column=2, pady=10, padx=20, sticky="n")

        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")
        self.optionmenu_1.set("Nº Inventário")
        self.optionmenu_2.set("Quant. de Cod")
        self.optionmenu_3.set("Carnes e Paes")
    
    file_path = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Layout\Inventários\Outros.csv"

    with open(file_path, mode='w', newline=''):
        pass

    def optionmenu_3_var_changed(self, *args):
        palavras_chave = ["Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"]
        selected_fornecedor = self.selected_fornecedor.get()

        if any(palavra in selected_fornecedor for palavra in palavras_chave):
            self.optionmenu_1.configure(state="disabled")
            self.optionmenu_2.configure(state="disabled")
            self.main_button_1.configure(state="disabled")
            self.entry.configure(state="disabled")
            if selected_fornecedor in ["Domart", "Forte Boi", "Suimartin", "Trigo Mais"]:
                self.radio_button_1.configure(state="disabled")
                self.radio_button_2.configure(state="disabled")
                self.radio_var.set(2)
            else:
                self.radio_button_1.configure(state="normal")
                self.radio_button_2.configure(state="normal")
                self.radio_var.set(0)
        else:
            pass

    def save_text(self):
        text = self.entry.get()
        file_path = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv"

        numbers = text.split()

        with open(file_path, mode='w', newline='') as file:
            writer = csv.writer(file)
            for number in numbers:
                writer.writerow([number])
                self.entry.delete(0, tkinter.END)

        is_outros_csv_empty = os.path.exists(file_path) and os.path.getsize(file_path) == 0

        if is_outros_csv_empty:
            pass
        else:
            self.optionmenu_2.configure(state="disabled")
            self.optionmenu_3.configure(state="disabled")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event_1(self):
        print("Inventário Iniciado")

        def abrir_inventário():
            pa.hotkey("win", "3")
            sleep(1)
            pa.hotkey("ctrl", "home")
            sleep(0.5)
            pa.press("alt")
            sleep(0.5)
            pa.press("e")
            sleep(0.5)
            pa.press("i")
            sleep(0.5)
            pa.press("f2")
            sleep(0.5)
            pa.press("tab")
            sleep(4)
            pa.press("down", presses=20)
            pa.press("up")
            pa.press("tab")

        abrir_inventário()

        palavras_chave = ["Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"]

        selected_fornecedor = self.optionmenu_3.get()

        print(selected_fornecedor)

        if any(palavra in selected_fornecedor for palavra in palavras_chave):
            try:
                selected_value_numero = self.optionmenu_1.get()
                selected_value_numero_int = int(selected_value_numero)
                print("Inventário Nº:", selected_value_numero_int)
            except:
                pass
        else:
            try:
                selected_value_numero = self.optionmenu_1.get()
                selected_value_numero_int = int(selected_value_numero)
                print("Inventário Nº:", selected_value_numero_int)
            except ValueError:
                print("Escolha um Número Valido!")
                pa.press("f9")
                sleep(1)
                print('________________________________')
                return

        selected_value = self.radio_var.get()

        selected_value_numero_invent = self.optionmenu_2.get()

        if selected_value_numero_invent == "Quant. de Cod":
            select_value_invent = "6"
        else:
            select_value_invent = selected_value_numero_invent
        
        selected_value_final = int(select_value_invent)

        if any(palavra in selected_fornecedor for palavra in palavras_chave):
            if selected_value == 0:
                print("Loja SMJ")
                pa.press("up", presses=6)
                pa.press("tab", presses=2)
                archive_name = f"{selected_fornecedor} SMJ AJUSTE ESTOQUE"
                pa.write(archive_name)
            elif selected_value == 1:
                print("Loja STT")
                pa.press("up", presses=6)
                pa.press("down")
                pa.press("tab", presses=2)
                archive_name = f"{selected_fornecedor} STT AJUSTE ESTOQUE"
                pa.write(archive_name)
            elif selected_value == 2:
                print("Loja VIX")
                pa.press("up", presses=6)
                pa.press("down", presses=2)
                pa.press("tab", presses=2)
                archive_name = f"{selected_fornecedor} VIX AJUSTE ESTOQUE"
                pa.write(archive_name)               
        else:
            if selected_value == 0:
                print("Loja SMJ")
                pa.press("up", presses=6)
                pa.press("tab", presses=2)
                archive_name = f"DIVERSOS SMJ {selected_value_numero} AJUSTE ESTOQUE"
                pa.write(archive_name)
            elif selected_value == 1:
                print("Loja STT")
                pa.press("up", presses=6)
                pa.press("down")
                pa.press("tab", presses=2)
                archive_name = f"DIVERSOS STT {selected_value_numero} AJUSTE ESTOQUE"
                pa.write(archive_name)
                print(archive_name)
            elif selected_value == 2:
                print("Loja VIX")
                pa.press("up", presses=6)
                pa.press("down", presses=2)
                pa.press("tab", presses=2)
                archive_name = f"DIVERSOS VIX {selected_value_numero} AJUSTE ESTOQUE"
                pa.write(archive_name)

        pa.click(41,37, interval=0.5)
        pa.click(20,844, interval=0.5)       
        sleep(1)

        file_path = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv"

        if selected_fornecedor in ("Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"):
            if selected_value == 0:
                caminho_csv = os.path.join("F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\Carnes e Paes", f"{selected_fornecedor} SMJ.csv")
            elif selected_value == 1:
                caminho_csv = os.path.join("F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\Carnes e Paes", f"{selected_fornecedor} STT.csv")
            elif selected_value == 2:
                caminho_csv = os.path.join("F:\\COMPRAS\\Automações\\Automação - Inventário\\Layout\\Inventários\\Carnes e Paes", f"{selected_fornecedor} VIX.csv")       
        else:
            if os.path.exists(file_path):
                if os.path.getsize(file_path) == 0:
                    if selected_value == 0:
                        caminho_csv = caminho_SMJ
                    elif selected_value == 1:
                        caminho_csv = caminho_STT
                    elif selected_value == 2:
                        caminho_csv = caminho_VIX
                else:
                    caminho_csv = r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv"
        
        contador = 0

        with open(caminho_csv) as arquivo:
            linhas = arquivo.readlines()
            sleep(2)

        for linha in linhas:

            if any(palavra in selected_fornecedor for palavra in palavras_chave):
                codigo = linha.split(';')[0]
                pa.write(codigo)
                pa.press("enter")
                sleep(0.1)
                pa.hotkey("alt", "o")
                sleep(0.5)
                pa.click(615, 845, duration=0.5)
                sleep(0.1)
                pa.click(711, 371, duration=0.5)

            elif caminho_csv == r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv":
                codigo = linha.split(';')[0]
                pa.write(codigo)
                pa.press("enter")
                sleep(0.1)
                pa.hotkey("alt", "o")
                sleep(0.5)
                pa.click(615, 845, duration=0.5)
                sleep(0.1)
                pa.click(711, 371, duration=0.5)
                    
            else:
                if contador >= selected_value_final:
                    break
                
                linha = linhas[contador]
                codigo = linha.split(';')[0]
                pa.write(codigo)
                pa.press("enter")
                sleep(0.1)
                pa.hotkey("alt", "o")
                sleep(0.5)
                pa.click(615, 845, duration=0.5)
                sleep(0.1)
                pa.click(711, 371, duration=0.5)

                contador += 1

        pa.press("esc")
        sleep(3)

        pa.click(1233,40, duration=1)
        sleep(1)

        pa.click(11,362)
        sleep(1)

        pa.hotkey("alt", "o")
        sleep(4)

        pa.click(461,48, duration=1)
        sleep(1)

        def is_capslock_on():
            return ctypes.windll.user32.GetKeyState(0x14) & 1 != 0

        if is_capslock_on():
            pa.press("capslock")
        else:
            pass

        archive_name = archive_name.upper()
        pa.write(archive_name)
        sleep(1)

        pa.press("enter")
        sleep(1)

        pa.press("f9", presses=2, interval=1)

        if selected_fornecedor in ("Segredos do Trigo", "Trigo Mais", "Trigo & Cia", "Domart", "Forte Boi", "Suimartin", "Granel"):
            pass
        elif caminho_csv == r"F:\COMPRAS\Automações\Automação - Inventário\Layout\Inventários\Outros.csv":
            with open(file_path, mode='w', newline=''):
                self.optionmenu_2.configure(state="normal")
                self.optionmenu_3.configure(state="normal")
                pass
        else:
            with open(caminho_csv, 'w') as arquivo:
                arquivo.writelines(linhas[contador:])
        
        if any(palavra in selected_fornecedor for palavra in palavras_chave):
            self.optionmenu_3.set("Carnes e Paes")
            self.optionmenu_1.configure(state="normal")
            self.optionmenu_2.configure(state="normal")
            self.main_button_1.configure(state="normal")
            self.entry.configure(state="normal")
            self.radio_button_1.configure(state="normal")
            self.radio_button_2.configure(state="normal")
        else:
            pass

        pa.press("esc")
        print('________________________________')
        
    def sidebar_button_event_2(self):
        self.destroy()

        file_path = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Inventários\Layout\Inventários\Outros.csv"

        with open(file_path, mode='w', newline=''):
            pass

        app = App()
        app.mainloop()

if __name__ == "__main__":
    app = App()
    app.mainloop()