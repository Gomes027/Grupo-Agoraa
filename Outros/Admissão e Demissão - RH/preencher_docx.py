import os
import shutil
import datetime
import customtkinter
from docx import Document
from docx2pdf import convert

customtkinter.set_appearance_mode("Dark")  # Modos: "System" (padrão), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Temas: "blue" (padrão), "green", "dark-blue"

class FormatadorDocumentos:
    def __init__(self, caminho_salvar_arquivo, diretorio_originais, referencias):
        self.caminho_salvar_arquivo = caminho_salvar_arquivo
        self.diretorio_originais = diretorio_originais
        self.referencias = referencias

    @staticmethod
    def _obter_mes_por_extenso(mes_numero):
        meses_em_portugues = {
            '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março',
            '04': 'Abril', '05': 'Maio', '06': 'Junho',
            '07': 'Julho', '08': 'Agosto', '09': 'Setembro',
            '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro'
        }
        return meses_em_portugues.get(mes_numero, 'Desconhecido')

    def formatar_documentos(self):
        data_atual = datetime.datetime.now()
        dia, mes_numero, ano = data_atual.strftime("%d"), data_atual.strftime("%m"), data_atual.strftime("%Y")
        mes_extenso = FormatadorDocumentos._obter_mes_por_extenso(mes_numero)

        # Atualizar referências com data atual
        self.referencias.update({"DD": dia, "MM": mes_extenso, "MNM": mes_numero, "AAAA": ano})

        nome_pasta = self.referencias["XXXX"]  # Exemplo de como o nome pode ser utilizado
        diretorio_editados = os.path.join(self.caminho_salvar_arquivo, nome_pasta)
        os.makedirs(diretorio_editados, exist_ok=True)

        diretorio_pdf = os.path.join(diretorio_editados, "PDF")
        diretorio_docx = os.path.join(diretorio_editados, "DOCX")
        os.makedirs(diretorio_pdf, exist_ok=True)
        os.makedirs(diretorio_docx, exist_ok=True)

        for arquivo in os.listdir(self.diretorio_originais):
            if arquivo.endswith(".docx"):
                caminho_arquivo_origem = os.path.join(self.diretorio_originais, arquivo)
                documento = Document(caminho_arquivo_origem)

                for paragrafo in documento.paragraphs:
                    for codigo, valor in self.referencias.items():
                        if codigo in paragrafo.text:
                            paragrafo.text = paragrafo.text.replace(codigo, valor)
                
                nome_arquivo_saida = f"{os.path.splitext(arquivo)[0]}.docx"
                caminho_arquivo_destino = os.path.join(diretorio_editados, nome_arquivo_saida)
                documento.save(caminho_arquivo_destino)
                print(f"Documento editado '{nome_arquivo_saida}' salvo com sucesso.")

                convert(caminho_arquivo_destino)

        # Organizar arquivos
        self._organizar_arquivos(diretorio_editados, diretorio_docx, diretorio_pdf)
        os.startfile(diretorio_editados)

    def _organizar_arquivos(self, diretorio_editados, diretorio_docx, diretorio_pdf):
        for arquivo in os.listdir(diretorio_editados):
            caminho_arquivo = os.path.join(diretorio_editados, arquivo)
            if arquivo.endswith(".docx"):
                shutil.move(caminho_arquivo, os.path.join(diretorio_docx, arquivo))
            elif arquivo.endswith(".pdf"):
                shutil.move(caminho_arquivo, os.path.join(diretorio_pdf, arquivo))
        
        print("Arquivos movidos para as pastas DOCX e PDF.")

class InterfaceUsuario(customtkinter.CTk):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.configurar_ui()

    def configurar_ui(self):
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
        
        # Botões na barra lateral
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, fg_color="ForestGreen", text="Admissão", font=("Poppins", 12, "bold"), command=lambda: self.app.sidebar_button_event_1())
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, fg_color="FireBrick", text="Demissão", font=("Poppins", 12, "bold"), command=lambda: self.app.sidebar_button_event_2())
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)

        # Configurações de tema
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Tema:", font=("Poppins", 12, "bold"), anchor="w")

        # Adicionar campos de entrada
        self.entradas = {}
        self.entradas['nome'] = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Nome...")
        self.entradas['nome'].grid(row=0, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        self.entradas['cpf'] = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="CPF...")
        self.entradas['cpf'].grid(row=1, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")
        self.entradas['cpf'].bind("<KeyRelease>", self.format_cpf)

        self.entradas['rg'] = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="RG...")
        self.entradas['rg'].grid(row=2, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        self.entradas['funcao'] = customtkinter.CTkEntry(self, font=("Poppins", 12, "bold"), placeholder_text="Função...")
        self.entradas['funcao'].grid(row=3, column=1, columnspan=2, padx=(20, 20), pady=(20, 20), sticky="ew")

        # Opção de Tema
        self.appearance_mode_label.grid(row=6, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionmenu = customtkinter.CTkOptionMenu(self.sidebar_frame, font=("Poppins", 12, "bold"), values=["Light", "Dark", "System"],
                                                                       command=self.app.change_appearance_mode_event)
        self.appearance_mode_optionmenu.grid(row=8, column=0, padx=20, pady=(10, 10))
        self.appearance_mode_optionmenu.set("Dark")  # Definir o tema padrão

    def coletar_dados(self):
        # Método para coletar dados dos campos de entrada
        dados = {chave: entrada.get() for chave, entrada in self.entradas.items()}
        return dados
    
    def format_cpf(self, event=None):
        cpf = self.entradas['cpf'].get().replace(".", "").replace("-", "")  # Remove pontos e traço
        novo_cpf = ""

        if len(cpf) > 11:
            cpf = cpf[:11]  # Limita o CPF a 11 caracteres

        # Adiciona ponto após o terceiro e sexto dígitos, e traço após o nono
        for i in range(len(cpf)):
            if i in [3, 6]:
                novo_cpf += "."
            elif i == 9:
                novo_cpf += "-"
            novo_cpf += cpf[i]

        # Atualiza o campo de entrada do CPF com o valor formatado
        self.entradas['cpf'].delete(0, "end")
        self.entradas['cpf'].insert(0, novo_cpf)

    def limpar_campos(self):
        """Limpa os valores dos campos de entrada."""
        for entrada in self.entradas.values():
            entrada.delete(0, 'end')

class App:
    def __init__(self):
        self.ui = InterfaceUsuario(self)

    def run(self):
        self.ui.mainloop()

    def formatar_e_salvar_documentos(self, caminho_salvar_arquivo, diretorio_originais):
        dados = self.ui.coletar_dados()
        dia, mes_numero, ano = datetime.datetime.now().strftime("%d"), datetime.datetime.now().strftime("%m"), datetime.datetime.now().strftime("%Y")
        mes_nome = FormatadorDocumentos._obter_mes_por_extenso(mes_numero)
        
        referencias = {
            "XXXX": dados['nome'],
            "YYYY": dados['funcao'],
            "ZZZZ": dados['cpf'],
            "DD": dia,
            "MM": mes_nome,
            "MNM": mes_numero,
            "AAAA": ano,
            "WWWW": dados['rg'],
        }
        
        formatador = FormatadorDocumentos(caminho_salvar_arquivo, diretorio_originais, referencias)
        formatador.formatar_documentos()

    def sidebar_button_event_1(self):
        caminho_salvar_arquivo = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Admissão e Demissão - RH\Formulários\Admissão"
        diretorio_originais = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Admissão e Demissão - RH\Outros\Admissão_Docx_Originais"
        self.formatar_e_salvar_documentos(caminho_salvar_arquivo, diretorio_originais)
        self.ui.limpar_campos()

    def sidebar_button_event_2(self):
        caminho_salvar_arquivo = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Admissão e Demissão - RH\Formulários\Demissão"
        diretorio_originais = r"F:\COMPRAS\Automações.Compras\Automações\Outros\Admissão e Demissão - RH\Outros\Demissão_Docx_Originais"
        self.formatar_e_salvar_documentos(caminho_salvar_arquivo, diretorio_originais)
        self.ui.limpar_campos()

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

if __name__ == "__main__":
    app = App()
    app.run()