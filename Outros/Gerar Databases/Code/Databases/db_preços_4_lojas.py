import sqlite3
import pandas as pd

class ManipuladorPreços4Lojas:
    def __init__(self, caminho_planilha, nome_tabela, caminho_db):
        self.caminho_planilha = caminho_planilha
        self.nome_tabela = nome_tabela
        self.caminho_db = caminho_db
    
    def criar_tabela(self):
        conexao = sqlite3.connect(self.caminho_db)
        cursor = conexao.cursor()
        cursor.execute(f'DROP TABLE IF EXISTS {self.nome_tabela};')
        cursor.execute(f'''
            CREATE TABLE {self.nome_tabela} (
                codigo INTEGER,
                nome TEXT,
                nome_loja TEXT,
                loja INTEGER,
                custo_liquido FLOAT,
                preco_venda FLOAT,
                preco_promocao FLOAT,
                preco_venda2 FLOAT
            );
        ''')
        conexao.commit()
        conexao.close()
    
    def inserir_dados(self):
        df = pd.read_excel(self.caminho_planilha)
        dados = df.values.tolist()
        conexao = sqlite3.connect(self.caminho_db)
        cursor = conexao.cursor()
        cursor.executemany(f'''
            INSERT INTO {self.nome_tabela} (
                codigo, nome, nome_loja, loja, custo_liquido, 
                preco_venda, preco_promocao, preco_venda2
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        ''', dados)
        conexao.commit()
        conexao.close()
        print('Dados inseridos com sucesso no banco de dados SQLite relatorio_precos_4_lojas.')

if __name__ == "__main__":
    from config import *
    
    manipulador_preços_4_lojas = ManipuladorPreços4Lojas(dir_xlsx_preço_4_lojas, table_name_preço_4_lojas, dir_db_preços_4_lojas)
    manipulador_preços_4_lojas.criar_tabela()
    manipulador_preços_4_lojas.inserir_dados()