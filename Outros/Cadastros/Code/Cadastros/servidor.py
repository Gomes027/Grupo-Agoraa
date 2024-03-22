import tkinter as tk
import sqlite3
from tkinter import ttk
import pyperclip
from tkinter import messagebox
import threading
import socket
import json
import queue

class DataBaseHandler:
    def __init__(self, db_path):
        self.db_path = db_path

    def excluir_produto(self, ean):
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM dados_cadastros WHERE EAN = ?", (ean,))
            conn.commit()
        except sqlite3.Error as e:
            print(f"Erro ao excluir produto com EAN {ean}: {e}")
        finally:
            conn.close()

    def tem_dados(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT EXISTS(SELECT 1 FROM dados_cadastros LIMIT 1)")
        resultado = cursor.fetchone()[0]
        conn.close()
        return resultado

    def inicializar(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS dados_cadastros (
                PRIORIDADE INTEGER,
                EAN TEXT,
                DESCRICAO_COMPLETA TEXT,
                DESCRICAO_CONSUMIDOR TEXT,
                FORNECEDOR TEXT,
                MARCA TEXT,
                GRAMATURA TEXT,
                EMBALAGEM_COMPRA TEXT,
                NCM TEXT,
                SETOR TEXT,
                GRUPO TEXT,
                SUBGRUPO TEXT,
                CATEGORIA TEXT,
                MIX TEXT,
                TR TEXT,
                AP TEXT
            )
        ''')
        conn.commit()
        conn.close()

    def inserir_dados(self, data):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO dados_cadastros (PRIORIDADE, EAN, DESCRICAO_COMPLETA, DESCRICAO_CONSUMIDOR, FORNECEDOR, MARCA, GRAMATURA, EMBALAGEM_COMPRA, NCM, SETOR, GRUPO, SUBGRUPO, CATEGORIA, MIX, TR, AP)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', data)
        conn.commit()
        conn.close()

    def carregar_dados(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM dados_cadastros ORDER BY PRIORIDADE DESC')  # Ordenando pela coluna PRIORIDADE em ordem decrescente
        rows = cursor.fetchall()
        conn.close()
        return rows

class Server:
    def __init__(self, host, port, database_handler, update_queue):
        self.host = host
        self.port = port
        self.database_handler = database_handler
        self.update_queue = update_queue

    def handle_client_connection(self, client_socket):
        while True:
            try:
                request = client_socket.recv(1024).decode('utf-8')
                if not request:
                    break

                print(f"Recebido: {request}")

                if request.strip():
                    request_data = json.loads(request)
                    self.update_queue.put(request_data['data'])
                    self.database_handler.inserir_dados(tuple(request_data['data']))

                    response_data = {"status": "success", "message": "Dados recebidos com sucesso!"}
                    client_socket.sendall(json.dumps(response_data).encode('utf-8'))
                else:
                    print("Dados recebidos vazios ou inválidos.")
            except Exception as e:
                print(f"Erro: {e}")
                break

        client_socket.close()

    def start(self):
        server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        server_socket.bind((self.host, self.port))
        server_socket.listen(5)
        print(f"Servidor iniciado e ouvindo em {self.host}:{self.port}")

        try:
            while True:
                client_sock, address = server_socket.accept()
                print(f"Aceita conexão de {address}")
                client_thread = threading.Thread(target=self.handle_client_connection, args=(client_sock,))
                client_thread.start()
        finally:
            server_socket.close()

class Application:
    def __init__(self, database_handler, update_queue):
        self.database_handler = database_handler
        self.update_queue = update_queue
        self.janela = tk.Tk()
        self.tabela = None
        self.server = Server('127.0.0.1', 65432, database_handler, update_queue)
        self.criar_janela_com_tabela()

    def atualizar_visibilidade_janela(self):
        if self.database_handler.tem_dados():
            if not self.janela.winfo_viewable():
                self.janela.deiconify()
        else:
            if self.janela.winfo_viewable():
                self.janela.withdraw()

    def criar_janela_com_tabela(self):
        self.janela.title("CAD-Produtos")
        
        self.janela.geometry("800x600")

        if not self.database_handler.tem_dados():
            self.janela.withdraw()

        frame = tk.Frame(self.janela)
        frame.pack(fill=tk.BOTH, expand=True)

        # Definindo o estilo da tabela
        style = ttk.Style()
        style.theme_use("default")  # Você pode experimentar outros temas aqui
        style.configure("Treeview", background="#D3D3D3", foreground="black", rowheight=25, fieldbackground="#D3D3D3")
        style.map("Treeview", background=[('selected', '#347083')])

        # Definindo a cor do cabeçalho
        style.configure("Treeview.Heading", background="DarkGray", foreground="black")

        colunas = ("PRIORIDADE", "EAN", "DESCRICAO_COMPLETA", "DESCRICAO_CONSUMIDOR", "FORNECEDOR", "MARCA", "GRAMATURA", "EMBALAGEM_COMPRA", "NCM", "SETOR", "GRUPO", "SUBGRUPO", "CATEGORIA", "MIX", "TR", "AP")
        self.tabela = ttk.Treeview(frame, columns=colunas, show="headings")
        for col in colunas:
            self.tabela.heading(col, text=col.replace("_", " "), anchor="center")
            if "DESCRICAO" in col:
                self.tabela.column(col, anchor="w")
            else:
                self.tabela.column(col, anchor="center")

        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tabela.yview)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tabela.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=self.tabela.xview)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tabela.configure(xscrollcommand=scrollbar_x.set)

        self.tabela.pack(fill=tk.BOTH, expand=True)

        def update_tabela_from_server():
            while True:
                data = self.update_queue.get()
                if data:
                    if not self.janela.winfo_viewable():
                        self.janela.deiconify()
                    self.tabela.insert('', tk.END, values=data)

        def mover_horizontalmente(event):
            if event.keysym == 'Right':
                self.tabela.xview_scroll(80, 'units')
            elif event.keysym == 'Left':
                self.tabela.xview_scroll(-80, 'units')

        # Vincula eventos de teclado para mover horizontalmente
        self.janela.bind("<Right>", mover_horizontalmente)
        self.janela.bind("<Left>", mover_horizontalmente)

        def copiar_valor_celula(event):
            row_id = self.tabela.identify_row(event.y)
            col_id = self.tabela.identify_column(event.x)
            if row_id and col_id:
                col_id_int = int(col_id.replace('#', '')) - 1
                valor_celula = self.tabela.item(row_id)['values'][col_id_int]
                pyperclip.copy(valor_celula)
                print("Dados copiados para a área de transferência:", valor_celula)

        self.tabela.bind("<Button-1>", copiar_valor_celula)

        def perguntar_e_excluir_produto(event):
            row_id = self.tabela.identify_row(event.y)
            if row_id:
                item = self.tabela.item(row_id)
                valores = item['values']
                ean = valores[1]
                if messagebox.askyesno("Confirmar", "O cadastro já foi feito? Deseja excluir este produto da fila?"):
                    self.tabela.delete(row_id)
                    self.database_handler.excluir_produto(ean)
                    self.atualizar_visibilidade_janela()

        self.tabela.bind("<Double-1>", perguntar_e_excluir_produto)

        def carregar_dados_do_banco_de_dados():
            rows = self.database_handler.carregar_dados()
            for row in rows:
                self.tabela.insert('', tk.END, values=row)

        carregar_dados_do_banco_de_dados()

        server_thread = threading.Thread(target=self.server.start, daemon=True)
        server_thread.start()

        tabela_thread = threading.Thread(target=update_tabela_from_server, daemon=True)
        tabela_thread.start()

        self.janela.mainloop()

if __name__ == "__main__":
    db_handler = DataBaseHandler(r'DataBase\registros_servidor.db')
    db_handler.inicializar()
    update_queue = queue.Queue()  # Definindo a fila de atualização aqui
    app = Application(db_handler, update_queue)