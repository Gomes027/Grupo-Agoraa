import pandas as pd
import sqlite3

# Caminho da planilha Excel
caminho_planilha = r'F:\BI\Bases\relatorio_precos_4_lojas.xlsx'

# Nome da tabela no banco de dados SQLite
nome_tabela = 'precos_lojas'

# Função para criar a tabela no banco de dados SQLite
def criar_tabela(cursor):
    # Apaga os dados da tabela se ela já existir
    cursor.execute('''
        DROP TABLE IF EXISTS {};
    '''.format(nome_tabela))

    # Cria a tabela novamente
    cursor.execute('''
        CREATE TABLE {} (
            codigo INTEGER,
            nome TEXT,
            nome_loja TEXT,
            loja INTEGER,
            custo_liquido FLOAT,
            preco_venda FLOAT,
            preco_promocao FLOAT,
            preco_venda2 FLOAT
        )
    '''.format(nome_tabela))

# Função para inserir os dados da planilha no banco de dados SQLite
def inserir_dados(cursor, dados):
    cursor.executemany('''
        INSERT INTO {} (
            codigo, nome, nome_loja, loja, custo_liquido, 
            preco_venda, preco_promocao, preco_venda2
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    '''.format(nome_tabela), dados)


df = pd.read_excel(caminho_planilha)

# Conectar ao banco de dados SQLite
conexao = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_precos_4_lojas.db')
cursor = conexao.cursor()

# Criar a tabela no banco de dados SQLite
criar_tabela(cursor)

# Inserir os dados da planilha no banco de dados SQLite
dados = df.values.tolist()
inserir_dados(cursor, dados)

# Commit e fechar conexão
conexao.commit()
conexao.close()

print('Dados inseridos com sucesso no banco de dados SQLite.')