import customtkinter as ctk

# Define o modo de aparência para 'dark-blue'
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# Inicializa a janela principal
app = ctk.CTk()
app.title('')
app.geometry('1000x750')

# Ajusta a configuração para acomodar 4 linhas em vez de 3
for i in range(4):  # Agora temos 4 linhas
    app.grid_rowconfigure(i, weight=1)
for i in range(4):
    app.grid_columnconfigure(i, weight=1)

frame_a = ctk.CTkFrame(master=app, corner_radius=10)
frame_a.grid(row=1, column=0, rowspan=4, sticky="nsew", padx=(20, 10), pady=20)

frame_b1 = ctk.CTkFrame(master=app, corner_radius=10)
frame_b1.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=10, pady=(20, 10))

frame_b2 = ctk.CTkFrame(master=app, corner_radius=10)
frame_b2.grid(row=1, column=1, columnspan=2, sticky="nsew", padx=10, pady=(10, 10))

frame_b3 = ctk.CTkFrame(master=app, corner_radius=10)
frame_b3.grid(row=2, column=1, columnspan=2, sticky="nsew", padx=10, pady=(10, 10))

frame_b4 = ctk.CTkFrame(master=app, corner_radius=10)
frame_b4.grid(row=3, column=1, columnspan=3, sticky="nsew", padx=10, pady=(10, 20))

frame_c = ctk.CTkFrame(master=app, corner_radius=10)
frame_c.grid(row=0, column=3, rowspan=3, sticky="nsew", padx=(10, 20), pady=20)

# Label with text "Códigos:" inside frame_a
label_codes = ctk.CTkLabel(master=frame_c, text="Códigos:", anchor="w")  # Adjust font as needed
label_codes.pack(fill='x', padx=20, pady=(5, 0))

# Textbox for multi-line text input inside frame_a
textbox_codes = ctk.CTkTextbox(master=frame_c)
textbox_codes.pack(fill='both', expand=True, padx=20, pady=(0, 20))

# Label with text "Códigos:" inside frame_a
label_codes = ctk.CTkLabel(master=frame_b1, text="Nome do Arquivo:", anchor="w")  # Adjust font as needed
label_codes.pack(fill='x', padx=20, pady=(5, 0))

# Textbox for multi-line text input inside frame_a
textbox_codes = ctk.CTkTextbox(master=frame_b1)
textbox_codes.pack(fill='both', expand=True, padx=20, pady=(0, 20))

# Executa a aplicação
app.mainloop()