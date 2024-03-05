import os
import shutil
import datetime
import customtkinter
from docx import Document
from docx2pdf import convert

customtkinter.set_appearance_mode("Dark")  # Modos: "System" (padrão), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Temas: "blue" (padrão), "green", "dark-blue"

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.title(" ")
        self.geometry(f"{500}x{340}")

        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Opções", font=customtkinter.CTkFont(family="Poppins", size=18, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, fg_color="ForestGreen", text="Admissão", font=("Poppins", 12, "bold"), command=self.sidebar_button_event_1)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, fg_color="FireBrick", text="Demissão", font=("Poppins", 12, "bold"), command=self.sidebar_button_event_2)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Tema:", font=("Poppins", 12, "bold"), anchor="w")
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, font=("Poppins", 12, "bold"), values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 10))


        self.entry_1 = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Nome...")
        self.entry_1.grid(row=0, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        def format_cpf(entry):
            current_text = entry.get()
            current_text = ''.join(filter(str.isdigit, current_text))

            if len(current_text) > 11:
                current_text = current_text[:11]

            if len(current_text) <= 3:
                formatted_text = current_text
            elif len(current_text) <= 6:
                formatted_text = f"{current_text[:3]}.{current_text[3:]}"
            elif len(current_text) <= 9:
                formatted_text = f"{current_text[:3]}.{current_text[3:6]}.{current_text[6:]}"
            else:
                formatted_text = f"{current_text[:3]}.{current_text[3:6]}.{current_text[6:9]}-{current_text[9:]}"

            entry.delete(0, "end")
            entry.insert(0, formatted_text)
            cpf = formatted_text

        self.entry_2 = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="CPF...")
        self.entry_2.grid(row=1, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")
        self.entry_2.bind("<KeyRelease>", lambda event: format_cpf(self.entry_2))

        self.entry_3 = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="RG...")
        self.entry_3.grid(row=2, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        self.entry_4 = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Função...")
        self.entry_4.grid(row=3, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        self.appearance_mode_optionemenu.set("Dark")
    
    def formatar_documentos(self, caminho_salvar_arquivo, diretorio_originais):

        data_atual = datetime.datetime.now()

        dia = data_atual.strftime("%d")
        mes_numero = data_atual.strftime("%m")
        ano = data_atual.strftime("%Y")

        meses_em_portugues = {
            '01': 'Janeiro',
            '02': 'Fevereiro',
            '03': 'Março',
            '04': 'Abril',
            '05': 'Maio',
            '06': 'Junho',
            '07': 'Julho',
            '08': 'Agosto',
            '09': 'Setembro',
            '10': 'Outubro',
            '11': 'Novembro',
            '12': 'Dezembro'
        }

        mes_nome = meses_em_portugues.get(mes_numero, 'Desconhecido')

        nome = self.entry_1.get()
        CPF = self.entry_2.get()
        RG = self.entry_3.get()
        funcao = self.entry_4.get()

        referências = {
            "XXXX" : nome,
            "YYYY" : funcao,
            "ZZZZ" : CPF,
            "DD" : dia,
            "MM" : mes_nome,
            "MNM" : mes_numero,
            "AAAA" : ano,
            "WWWW" : RG,
        }

        nome_pasta = f"{nome}"
        diretorio_editados = os.path.join(caminho_salvar_arquivo, nome_pasta)
        os.makedirs(diretorio_editados, exist_ok=True)

        diretorio_pdf = os.path.join(diretorio_editados, "PDF")
        diretorio_docx = os.path.join(diretorio_editados, "DOCX")
        os.makedirs(diretorio_pdf, exist_ok=True)
        os.makedirs(diretorio_docx, exist_ok=True)

        for arquivo in os.listdir(diretorio_originais):
            if arquivo.endswith(".docx"):
                nome_arquivo = os.path.splitext(arquivo)[0]
                
                caminho_arquivo_origem = os.path.join(diretorio_originais, arquivo)
                documento = Document(caminho_arquivo_origem)
                
                for paragrafo in documento.paragraphs:
                    for codigo, valor in referências.items():
                        paragrafo.text = paragrafo.text.replace(codigo, valor)
                
                nome_arquivo_saida = f"{nome_arquivo}.docx"
                
                caminho_arquivo_destino = os.path.join(diretorio_editados, nome_arquivo_saida)
                documento.save(caminho_arquivo_destino)
                print(f"Documento editado '{nome_arquivo_saida}' salvo com sucesso.")

                convert(caminho_arquivo_destino)

        for arquivo in os.listdir(diretorio_editados):
            caminho_arquivo = os.path.join(diretorio_editados, arquivo)
            if arquivo.endswith(".docx"):
                shutil.move(caminho_arquivo, os.path.join(diretorio_docx, arquivo))
            elif arquivo.endswith(".pdf"):
                shutil.move(caminho_arquivo, os.path.join(diretorio_pdf, arquivo))

        print("Arquivos movidos para as pastas DOCX e PDF.")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def sidebar_button_event_1(self):
        caminho_salvar_arquivo = r"F:\COMPRAS\Automações\Automação - RH\Formulários\Admissão"
        diretorio_originais = r"F:\COMPRAS\Automações\Automação - RH\Formulários\Admissão\Originais"

        self.formatar_documentos(caminho_salvar_arquivo, diretorio_originais)

        self.destroy()
        app = App()
        app.mainloop()
        
    def sidebar_button_event_2(self):
        caminho_salvar_arquivo = r"F:\COMPRAS\Automações\Automação - RH\Formulários\Demissão"
        diretorio_originais = r"F:\COMPRAS\Automações\Automação - RH\Formulários\Demissão\Originais"

        self.formatar_documentos(caminho_salvar_arquivo, diretorio_originais)
        
        self.destroy()
        app = App()
        app.mainloop()

if __name__ == "__main__":
    app = App()
    app.mainloop()