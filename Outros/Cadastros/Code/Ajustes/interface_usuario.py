import customtkinter
from tkinter import BooleanVar
from tkinter import messagebox
from Cadastros.banco_de_dados import BancoDeDados

class InterfaceUsuario(customtkinter.CTk):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.banco = BancoDeDados()
        self.configurar_ui()
        self.protocol("WM_DELETE_WINDOW", self.app.close)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def configurar_ui(self):
        self.title(" ")
        self.geometry(f"{1000}x{700}")  # Ajuste o tamanho conforme necessário para acomodar todos os elementos

        # Ajuste na configuração de layout para acomodar a sidebar e as colunas de entrada
        self.grid_columnconfigure(0, weight=0)  # Sidebar
        self.grid_columnconfigure((1, 2), weight=1)  # Conteúdo principal
        self.grid_rowconfigure(0, weight=1)

        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=10)
        self.sidebar_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure((0, 1, 2), weight=1)

        # Configurações de fontes e tamanhos
        self.fonte_padrao_1 = ("Poppins", 12)
        self.fonte_padrao_2 = ("Poppins", 13, "bold")
        self.fonte_padrao_3 = ("Poppins", 14, "bold")
        
        # Divisão da sidebar em frames para botões, MIX/TR e tema
        self.buttons_frame = customtkinter.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.buttons_frame.grid(row=0, column=0, sticky="nsew")
        self.buttons_frame.grid_rowconfigure((0, 1), weight=1)

        self.mix_tr_frame = customtkinter.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.mix_tr_frame.grid(row=1, column=0, sticky="nsew")
        self.mix_tr_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 ,13 , 14), weight=1)

        self.theme_frame = customtkinter.CTkFrame(self.sidebar_frame, fg_color="transparent")
        self.theme_frame.grid(row=2, column=0, sticky="nsew")
        self.theme_frame.grid_columnconfigure(0, weight=1)
        self.theme_frame.grid_rowconfigure((0, 1, 3, 4), weight=1)
        
        # Botões na barra lateral
        self.sidebar_button_1 = customtkinter.CTkButton(self.buttons_frame, fg_color="ForestGreen", text="SALVAR", font=self.fonte_padrao_3, command=self.salvar_botao)
        self.sidebar_button_1.grid(row=0, column=0, padx=20, pady=(20, 0))
        self.sidebar_button_2 = customtkinter.CTkButton(self.buttons_frame, fg_color="FireBrick", text="EXPORTAR", font=self.fonte_padrao_3, command=self.exportar_botao)
        self.sidebar_button_2.grid(row=1, column=0, padx=20, pady=(0, 20))

        # Adicionando os checkboxes para MIX na sidebar
        self.sidebar_label_mix = customtkinter.CTkLabel(self.mix_tr_frame, text="MIX:", font=self.fonte_padrao_2)
        self.sidebar_label_mix.grid(row=0, column=0, padx=20, pady=(10, 0), sticky="w")

        mix_opcoes = ["SMJ", "STT", "VIX", "MCP"]
        self.checkboxes_mix = {}
        for i, opcao in enumerate(mix_opcoes):
            self.checkboxes_mix[opcao] = customtkinter.CTkCheckBox(self.mix_tr_frame, text=opcao)
            self.checkboxes_mix[opcao].grid(row=1+i, column=0, padx=20, pady=1, sticky="w")

        # Adicionando os checkboxes para TR na sidebar
        self.sidebar_label_tr = customtkinter.CTkLabel(self.mix_tr_frame, text="TR:", font=self.fonte_padrao_2)
        self.sidebar_label_tr.grid(row=5, column=0, padx=20, pady=(10, 0), sticky="w")

        tr_opcoes = ["SMJ", "STT", "VIX", "MCP"]
        self.checkboxes_tr = {}
        for i, opcao in enumerate(tr_opcoes):
            self.checkboxes_tr[opcao] = customtkinter.CTkCheckBox(self.mix_tr_frame, text=opcao)
            self.checkboxes_tr[opcao].grid(row=6+i, column=0, padx=20, pady=1, sticky="w")

        # Adicionando a checkbox para AP na sidebar, dentro do self.mix_tr_frame
        self.ap_var = BooleanVar(value=True)
        self.sidebar_label_ap = customtkinter.CTkLabel(self.mix_tr_frame, text="AP:", font=self.fonte_padrao_2)
        self.sidebar_label_ap.grid(row=6+len(tr_opcoes)+1, column=0, padx=20, pady=(10, 0), sticky="w")
        self.checkbox_ap = customtkinter.CTkCheckBox(self.mix_tr_frame, text="SIM", variable=self.ap_var)
        self.checkbox_ap.grid(row=6+len(tr_opcoes)+2, column=0, padx=20, pady=(0, 10), sticky="w")

        # Adicionando a checkbox para PRIORIDADE na sidebar, dentro do self.mix_tr_frame
        self.prioridade_var = BooleanVar(value=False)
        self.sidebar_label_prioridade = customtkinter.CTkLabel(self.mix_tr_frame, text="Prioridade:", font=self.fonte_padrao_2)
        self.sidebar_label_prioridade.grid(row=6+len(tr_opcoes)+3, column=0, padx=20, pady=(10, 0), sticky="w")
        self.checkbox_prioridade = customtkinter.CTkCheckBox(self.mix_tr_frame, text="SIM", variable=self.prioridade_var)
        self.checkbox_prioridade.grid(row=6+len(tr_opcoes)+4, column=0, padx=20, pady=(0, 20), sticky="w")

        # Configurações do menu tema
        self.appearance_mode_label = customtkinter.CTkLabel(self.theme_frame, text="Tema:", font=self.fonte_padrao_2)
        self.appearance_mode_label.grid(row=1, column=0, padx=20, pady=(20, 0), sticky="w")
        self.appearance_mode_optionmenu = customtkinter.CTkOptionMenu(self.theme_frame, font=self.fonte_padrao_2, values=["Light", "Dark", "System"], command=self.change_appearance_mode_event)
        self.appearance_mode_optionmenu.grid(row=2, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.appearance_mode_optionmenu.set("Dark")  # Define o tema padrão

        # Configuração das colunas para os campos de entrada
        coluna_esquerda_frame = customtkinter.CTkFrame(self, corner_radius=10)
        coluna_esquerda_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        coluna_esquerda_frame.grid_columnconfigure(0, weight=1)
        coluna_esquerda_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), weight=1)

        coluna_direita_frame = customtkinter.CTkFrame(self, corner_radius=10)
        coluna_direita_frame.grid(row=0, column=2, padx=10, pady=10, sticky="nsew")
        coluna_direita_frame.grid_columnconfigure(0, weight=1)
        coluna_direita_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13), weight=1)
        
        # Criação de campos de entrada para informações do produto
        self.entradas = {}
        
        # Coluna Esquerda
        # Para MOTIVO_CADASTRO
        label_motivo_cadastro = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Motivo do Cadastro")
        label_motivo_cadastro.grid(row=0, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['MOTIVO_CADASTRO'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Descreva o motivo do cadastro... (Opcional)")
        self.entradas['MOTIVO_CADASTRO'].grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para EAN
        label_ean = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Código de Barras")
        label_ean.grid(row=2, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['EAN'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Insira o código de barras...")
        self.entradas['EAN'].grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 10))
        self.entradas['EAN'].bind("<KeyRelease>", self.format_ean)

        # Para DESCRICAO_COMPLETA
        label_descricao_completa = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Descrição Completa do Produto")
        label_descricao_completa.grid(row=4, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['DESCRICAO_COMPLETA'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Digite a descrição completa do produto...")
        self.entradas['DESCRICAO_COMPLETA'].grid(row=5, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para DESCRICAO_CONSUMIDOR
        label_descricao_consumidor = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Descrição do Consumidor")
        label_descricao_consumidor.grid(row=6, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['DESCRICAO_CONSUMIDOR'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Descreva o produto como visto pelo consumidor...")
        self.entradas['DESCRICAO_CONSUMIDOR'].grid(row=7, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para FORNECEDOR
        label_fornecedor = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Fornecedor")
        label_fornecedor.grid(row=8, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['FORNECEDOR'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Nome ou empresa do fornecedor...")
        self.entradas['FORNECEDOR'].grid(row=9, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para MARCA
        label_marca = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Marca")
        label_marca.grid(row=10, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['MARCA'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Indique a marca do produto... (Opcional)")
        self.entradas['MARCA'].grid(row=11, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para GRAMATURA
        label_gramatura = customtkinter.CTkLabel(coluna_esquerda_frame, font=self.fonte_padrao_2, text="Gramatura")
        label_gramatura.grid(row=12, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['GRAMATURA'] = customtkinter.CTkEntry(coluna_esquerda_frame, font=self.fonte_padrao_1, placeholder_text="Peso ou volume do produto (ex: 500g, 1L)...")
        self.entradas['GRAMATURA'].grid(row=13, column=0, sticky="nsew", padx=20, pady=(0, 30))

        # Coluna Direita
        # Para EMBALAGEM_COMPRA
        label_embalagem_compra = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Embalagem de Compra")
        label_embalagem_compra.grid(row=0, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['EMBALAGEM_COMPRA'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Quantidade de unidades em uma caixa...")
        self.entradas['EMBALAGEM_COMPRA'].grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 10))
        self.entradas['EMBALAGEM_COMPRA'].bind("<KeyRelease>", self.format_embalagem_compra)

        # Para NCM
        label_ncm = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="NCM")
        label_ncm.grid(row=2, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['NCM'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Número da Nomenclatura Comum do Mercosul...")
        self.entradas['NCM'].grid(row=3, column=0, sticky="nsew", padx=20, pady=(0, 10))
        self.entradas['NCM'].bind("<KeyRelease>", self.format_ncm)

        # Para CODIGO_INTERNO
        label_codigo_interno = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Código do Fornecedor")
        label_codigo_interno.grid(row=4, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['CODIGO_INTERNO'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Código interno usado pelo fornecedor... (Opcional)")
        self.entradas['CODIGO_INTERNO'].grid(row=5, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para SETOR
        label_setor = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Setor")
        label_setor.grid(row=6, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['SETOR'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Setor de alocação do produto... (Opcional)")
        self.entradas['SETOR'].grid(row=7, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para GRUPO
        label_grupo = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Grupo")
        label_grupo.grid(row=8, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['GRUPO'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Grupo do produto... (Opcional)")
        self.entradas['GRUPO'].grid(row=9, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para SUBGRUPO
        label_subgrupo = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Subgrupo")
        label_subgrupo.grid(row=10, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['SUBGRUPO'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Subgrupo do produto... (Opcional)")
        self.entradas['SUBGRUPO'].grid(row=11, column=0, sticky="nsew", padx=20, pady=(0, 10))

        # Para CATEGORIA
        label_categoria = customtkinter.CTkLabel(coluna_direita_frame, font=self.fonte_padrao_2, text="Categoria")
        label_categoria.grid(row=12, column=0, sticky="w", padx=20, pady=(10, 0))
        self.entradas['CATEGORIA'] = customtkinter.CTkEntry(coluna_direita_frame, font=self.fonte_padrao_1, placeholder_text="Categoria do produto... (Opcional)")
        self.entradas['CATEGORIA'].grid(row=13, column=0, sticky="nsew", padx=20, pady=(0, 30))

    def format_ean(self, event=None):
        ean = self.entradas['EAN'].get().replace(".", "").replace("-", "")
        novo_ean = ""

        if len(ean) > 13:
            ean = ean[:13]

        # Garante que apenas dígitos numéricos sejam incluídos
        for char in ean:
            if char.isdigit():
                novo_ean += char

        self.entradas['EAN'].delete(0, "end")
        self.entradas['EAN'].insert(0, novo_ean)

    def format_embalagem_compra(self, event=None):
        embalagem = self.entradas['EMBALAGEM_COMPRA'].get()
        nova_embalagem = ""

        # Garante que apenas dígitos numéricos sejam incluídos
        for char in embalagem:
            if char.isdigit():
                nova_embalagem += char

        self.entradas['EMBALAGEM_COMPRA'].delete(0, "end")
        self.entradas['EMBALAGEM_COMPRA'].insert(0, nova_embalagem)

    def format_ncm(self, event=None):
        ncm = self.entradas['NCM'].get()
        novo_ncm = ""

        if len(ncm) > 8:
            ncm = ncm[:8]  # Limita o NCM a 8 caracteres

        # Garante que apenas dígitos numéricos sejam incluídos
        for char in ncm:
            if char.isdigit():
                novo_ncm += char

        self.entradas['NCM'].delete(0, "end")
        self.entradas['NCM'].insert(0, novo_ncm)

    def coletar_dados(self):
        # Método para coletar dados dos campos de entrada
        dados = {chave: entrada.get() for chave, entrada in self.entradas.items()}
        return dados

    def salvar_botao(self):
        dados = self.coletar_dados()
        campos_faltantes = []

        # Verificar se as entradas obrigatórias estão preenchidas (exceto as exceções mencionadas)
        for campo in ['EAN', 'DESCRICAO_COMPLETA', 'DESCRICAO_CONSUMIDOR', 'FORNECEDOR', 'GRAMATURA', 'EMBALAGEM_COMPRA', 'NCM']:
            if not dados.get(campo):
                campos_faltantes.append(campo)

        # Verifica se pelo menos um MIX está selecionado
        mix_selecionado = any(self.checkboxes_mix[mix].get() for mix in ["SMJ", "STT", "VIX", "MCP"])
        if not mix_selecionado:
            campos_faltantes.append("MIX (pelo menos um)")

        if campos_faltantes:
            # Se existem campos faltantes, mostra uma mensagem de erro e interrompe a operação
            messagebox.showerror("Campos Obrigatórios Faltando", f"Por favor, preencha todos os campos obrigatórios:\n\n{', '.join(campos_faltantes)}")
            return
        
        ean = dados.get('EAN')
        ean_existente, codigo, nome = self.banco.verificar_ean_existente(ean)
        
        if ean_existente:
            messagebox.showwarning("Produto Já Cadastrado", f"O produto com EAN {ean} já está cadastrado no sistema.\nProduto: {nome}\n\nCódigo interno: {codigo}")
            return
        
        # Inicializar campos dos checkboxes com vazio
        for mix in ["SMJ", "STT", "VIX", "MCP"]:
            dados[f"MIX {mix}"] = "SIM" if self.checkboxes_mix[mix].get() else ""
        for tr in ["SMJ", "STT", "VIX", "MCP"]:
            dados[f"TR {tr}"] = "SIM" if self.checkboxes_tr[tr].get() else ""
        # AP é único, então não precisa de loop
        dados['AP'] = "SIM" if self.checkbox_ap.get() else ""
        dados['PRIORIDADE'] = "SIM" if self.checkbox_prioridade.get() else ""
        
        sucesso = self.banco.inserir_dados_temp(dados)
        descricao_consumidor = self.entradas['DESCRICAO_CONSUMIDOR'].get()

        if sucesso:
            messagebox.showinfo("Sucesso", f"Produto {descricao_consumidor} salvo com sucesso!")
        else:
            messagebox.showerror("Erro", f"O Produto {descricao_consumidor} não foi salvo corretamente.")

    def exportar_botao(self):
        nome_fornecedor = self.entradas['FORNECEDOR'].get().strip()
        nome_do_arquivo = f"CADASTRO {nome_fornecedor}.xlsx"

        sucesso = self.banco.exportar_arquivo(nome_do_arquivo)
        
        if sucesso:
            self.limpar_campos()
            messagebox.showinfo("Sucesso", "Dados exportados com sucesso!")
        elif sucesso == 0:
            messagebox.showerror("Erro", "Nenhum dado para exportar.")
        else:
            messagebox.showerror("Erro", "Falha ao exportar os dados.")

    def limpar_campos(self):
        """Limpa os valores dos campos de entrada somente se estiverem preenchidos."""
        for entrada in self.entradas.values():
            if entrada.get():  # Verifica se há conteúdo na entrada
                entrada.delete(0, 'end')  # Limpa o conteúdo se a entrada não estiver vazia