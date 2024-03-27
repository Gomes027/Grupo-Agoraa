import pandas as pd
import sqlite3

# Caminho da planilha Excel
caminho_planilha = r'F:\COMPRAS\relatorio_tresmann.xlsx'

# Nome da tabela no banco de dados SQLite
nome_tabela = 'dados_compras'

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
            cod_barras INTEGER,
            loja TEXT,
            fora_mix TEXT,
            fornecedor TEXT,
            nome_fantasia TEXT,
            peso TEXT,
            setor TEXT,
            grupo TEXT,
            subgrupo TEXT,
            qtde_embalagem INTEGER,
            estoque_minimo INTEGER,
            venda_ultimos_30_dias FLOAT,
            venda_ultimos_12_meses FLOAT,
            venda_ultimos_7_dias FLOAT,
            venda_dia FLOAT,
            ultimo_custo FLOAT,
            media_venda_ultimos_30_dias FLOAT,
            estoque_atual FLOAT,
            pedido_aberto FLOAT,
            validade TEXT,
            qtde_perda FLOAT,
            tipo_promocao TEXT,
            marca TEXT,
            qtde_troca FLOAT,
            flag1 TEXT,
            flag2 TEXT,
            flag3 TEXT,
            flag4 TEXT,
            flag5 TEXT,
            vcto_medio FLOAT,
            tipo_emp_fornecedor TEXT,
            estoque_rede FLOAT,
            emb_transf INTEGER,
            qtde_promocao FLOAT,
            categoria TEXT,
            qtde_sem_entrega INTEGER
        );
    '''.format(nome_tabela))

# Função para inserir ou atualizar os dados da planilha no banco de dados SQLite
def inserir_atualizar_dados(cursor, dados):
    cursor.executemany('''
        REPLACE INTO {} (
            codigo, nome, cod_barras, loja, fora_mix, fornecedor, nome_fantasia,
            peso, setor, grupo, subgrupo, qtde_embalagem, estoque_minimo,
            venda_ultimos_30_dias, venda_ultimos_12_meses, venda_ultimos_7_dias,
            venda_dia, ultimo_custo, media_venda_ultimos_30_dias, estoque_atual,
            pedido_aberto, validade, qtde_perda, tipo_promocao, marca, qtde_troca,
            flag1, flag2, flag3, flag4, flag5, vcto_medio, tipo_emp_fornecedor,
            estoque_rede, emb_transf, qtde_promocao, categoria, qtde_sem_entrega
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    '''.format(nome_tabela), dados)

df = pd.read_excel(caminho_planilha)

# Conectar ao banco de dados SQLite
conexao = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_tresmann.db')
cursor = conexao.cursor()

# Criar a tabela no banco de dados SQLite
criar_tabela(cursor)

# Inserir ou atualizar os dados da planilha no banco de dados SQLite
dados = df.values.tolist()
inserir_atualizar_dados(cursor, dados)

# Commit e fechar conexão
conexao.commit()
conexao.close()

print('Dados inseridos/atualizados com sucesso no banco de dados SQLite.')