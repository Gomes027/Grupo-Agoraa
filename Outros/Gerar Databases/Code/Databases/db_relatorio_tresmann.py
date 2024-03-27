import sqlite3
import pandas as pd

class ManipuladorDadosCompras:
    def __init__(self, caminho_planilha, nome_tabela, caminho_bd):
        self.caminho_planilha = caminho_planilha
        self.nome_tabela = nome_tabela
        self.caminho_bd = caminho_bd
    
    def criar_tabela(self):
        conexao = sqlite3.connect(self.caminho_bd)
        cursor = conexao.cursor()
        cursor.execute(f'DROP TABLE IF EXISTS {self.nome_tabela};')
        cursor.execute(f'''
            CREATE TABLE {self.nome_tabela} (
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
        ''')
        conexao.commit()
        conexao.close()
    
    def inserir_atualizar_dados(self):
        df = pd.read_excel(self.caminho_planilha)
        dados = df.values.tolist()
        conexao = sqlite3.connect(self.caminho_bd)
        cursor = conexao.cursor()
        cursor.executemany(f'''
            REPLACE INTO {self.nome_tabela} (
                codigo, nome, cod_barras, loja, fora_mix, fornecedor, nome_fantasia,
                peso, setor, grupo, subgrupo, qtde_embalagem, estoque_minimo,
                venda_ultimos_30_dias, venda_ultimos_12_meses, venda_ultimos_7_dias,
                venda_dia, ultimo_custo, media_venda_ultimos_30_dias, estoque_atual,
                pedido_aberto, validade, qtde_perda, tipo_promocao, marca, qtde_troca,
                flag1, flag2, flag3, flag4, flag5, vcto_medio, tipo_emp_fornecedor,
                estoque_rede, emb_transf, qtde_promocao, categoria, qtde_sem_entrega
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        ''', dados)
        conexao.commit()
        conexao.close()
        print('Dados inseridos/atualizados com sucesso no banco de dados SQLite relatorio_tresmann.')

if __name__ == "__main__":
    from config import *
    
    manipulador_dados_compras = ManipuladorDadosCompras(dir_xlsx_relatorio_tresmann, table_name_relatorio_tresmann, dir_db_relatorio_tresmann)
    manipulador_dados_compras.criar_tabela()
    manipulador_dados_compras.inserir_atualizar_dados()
