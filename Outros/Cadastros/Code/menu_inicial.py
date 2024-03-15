import customtkinter
from Cadastros.app_cadastro import AppCadastros

# Configurações iniciais da aparência e tema
customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")

class MenuInicial(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title(" ")
        self.geometry("300x300")  # Ajustar conforme necessário
        self.configurar_ui()

    def configurar_ui(self):
        self.frame_principal = customtkinter.CTkFrame(self)
        self.frame_principal.pack(pady=20, padx=20, fill="both", expand=True)

        # Configurar grid do frame_principal com colunas extras para espaçamento
        self.frame_principal.grid_rowconfigure((0, 5), weight=1)
        self.frame_principal.grid_rowconfigure((1, 2, 3, 4), weight=0)
        
        # Colunas para espaçamento lateral
        self.frame_principal.grid_columnconfigure(0, weight=1)
        self.frame_principal.grid_columnconfigure(1, weight=0)
        self.frame_principal.grid_columnconfigure(2, weight=1)

        button_width = 200  # Largura dos botões
        button_height = 40  # Altura dos botões

        # Botão para "AJUSTES NO MIX"
        self.btn_ajustes_mix = customtkinter.CTkButton(self.frame_principal, text="AJUSTES NO MIX", width=button_width, height=button_height, font=("Poppins", 12, "bold"), command=self.abrir_ajustes_mix)
        self.btn_ajustes_mix.grid(row=1, column=1, pady=10, padx=10)

        # Botão para "CADASTRO DE PRODUTOS"
        self.btn_cadastro_produtos = customtkinter.CTkButton(self.frame_principal, text="CADASTRO DE PRODUTOS", width=button_width, height=button_height, font=("Poppins", 12, "bold"), command=self.abrir_cadastro_produtos)
        self.btn_cadastro_produtos.grid(row=2, column=1, pady=10, padx=10)

        # Botão para "CADASTRO DE FORNECEDOR"
        self.btn_cadastro_fornecedor = customtkinter.CTkButton(self.frame_principal, text="CADASTRO DE FORNECEDOR", width=button_width, height=button_height, font=("Poppins", 12, "bold"), command=self.abrir_cadastro_fornecedor)
        self.btn_cadastro_fornecedor.grid(row=3, column=1, pady=10, padx=10)

        # Botão para "UNIFICAR CÓDIGOS"
        self.btn_unificar_codigos = customtkinter.CTkButton(self.frame_principal, text="UNIFICAR CÓDIGOS", width=button_width, height=button_height, font=("Poppins", 12, "bold"), command=self.abrir_unificar_codigos)
        self.btn_unificar_codigos.grid(row=4, column=1, pady=10, padx=10)

    def show(self):
        # Este método mostra o menu principal novamente
        self.deiconify()

    def abrir_ajustes_mix(self):
        pass

    def abrir_cadastro_produtos(self):
        # Esconde a janela do menu inicial
        self.withdraw()
        
        # Inicia a janela de cadastro de produtos
        app = AppCadastros(on_close=self.show)
        app.run()

    def abrir_cadastro_fornecedor(self):
        # Aqui você abriria a janela ou função correspondente
        print("Abrindo CADASTRO DE FORNECEDOR")

    def abrir_unificar_codigos(self):
        # Aqui você abriria a janela ou função correspondente
        print("Abrindo UNIFICAR CÓDIGOS")

if __name__ == "__main__":
    main_menu = MenuInicial()
    main_menu.mainloop()
