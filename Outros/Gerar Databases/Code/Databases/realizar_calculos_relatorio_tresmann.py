import time
import sqlite3

class AtualizadorDadosCompras:
    def __init__(self, caminho_bd_compras, caminho_bd_precos):
        self.caminho_bd_compras = caminho_bd_compras
        self.caminho_bd_precos = caminho_bd_precos
        self.conn = sqlite3.connect(self.caminho_bd_compras)
        self.cursor = self.conn.cursor()
    
    def adicionar_colunas(self, colunas_a_adicionar):
        for coluna in colunas_a_adicionar:
            self.cursor.execute("PRAGMA table_info(dados_compras)")
            colunas_existentes = [col[1] for col in self.cursor.fetchall()]
            if coluna.split()[0] not in colunas_existentes:
                self.cursor.execute(f'ALTER TABLE dados_compras ADD COLUMN {coluna}')
    
    def calcular_atualizar_valores(self):
        self.cursor.execute('''UPDATE dados_compras 
                              SET vda_media_30d_12m = venda_ultimos_12_meses / 12,
                                  maior_vda = CASE WHEN venda_ultimos_30_dias > (venda_ultimos_12_meses / 12) THEN venda_ultimos_30_dias ELSE (venda_ultimos_12_meses / 12) END''')

        self.cursor.execute('''UPDATE dados_compras 
                              SET venda_diaria = maior_vda / 30,
                                  dias_para_vencer = CASE
                                        WHEN (strftime('%s', validade) - strftime('%s', 'now', '-1 day')) / 86400 < 0 THEN 0
                                        ELSE CAST((strftime('%s', validade) - strftime('%s', 'now', '-1 day')) / 86400 AS INTEGER)
                                        END''')

        self.cursor.execute('''UPDATE dados_compras 
                              SET projecao_vendas = dias_para_vencer * venda_diaria''')

        self.cursor.execute('''UPDATE dados_compras                   
                                SET est_proj = CASE WHEN estoque_atual > projecao_vendas THEN 1 ELSE 0 END,
                                    est_excedente = CASE WHEN est_proj = 0 THEN 0 ELSE MAX(estoque_atual - projecao_vendas, 0) END''')

        precos_dict = {}
        for row in self.cursor.execute('SELECT codigo, loja FROM dados_compras'):
            codigo, loja = row
            precos = self.get_precos(codigo, loja)
            if precos:
                p_custo, p_venda = precos
                precos_dict[(codigo, loja)] = (p_custo, p_venda)

        for (codigo, loja), (p_custo, p_venda) in precos_dict.items():
            self.cursor.execute('''UPDATE dados_compras 
                                  SET p_custo = ?,
                                      p_venda = ?
                                  WHERE codigo = ? AND loja = ?''', (p_custo, p_venda, codigo, loja))

        self.cursor.execute('''UPDATE dados_compras 
                              SET total_custo = est_excedente * p_custo,
                                  total_venda = est_excedente * p_venda''')

        self.cursor.execute('''UPDATE dados_compras 
                              SET prazo = CASE WHEN dias_para_vencer < 60 THEN 'SIM' ELSE 'NÃO' END,
                                  est_atual_x_emb = CASE WHEN estoque_atual < qtde_embalagem THEN 'SIM' ELSE 'NÃO' END''')
    
    def get_precos(self, codigo, loja):
        precos_conn = sqlite3.connect(self.caminho_bd_precos)
        precos_cursor = precos_conn.cursor()
        precos_cursor.execute('''SELECT custo_liquido, preco_venda 
                                 FROM precos_lojas 
                                 WHERE codigo = ? AND nome_loja = ?''', (codigo, loja))
        row = precos_cursor.fetchone()
        precos_conn.close()
        return row
    
    def atualizar_dados_compras(self):
        try:
            start_time = time.time()
            colunas_a_adicionar = [
                'vda_media_30d_12m REAL',
                'maior_vda REAL',
                'venda_diaria REAL',
                'dias_para_vencer INTEGER',
                'projecao_vendas REAL',
                'est_proj INTEGER',
                'est_excedente REAL',
                'p_custo REAL',
                'p_venda REAL',
                'total_custo REAL',
                'total_venda REAL',
                'prazo TEXT',
                'est_atual_x_emb TEXT'
            ]
            self.adicionar_colunas(colunas_a_adicionar)
            self.calcular_atualizar_valores()
            self.conn.commit()
            end_time = time.time()
            execution_time_minutes = (end_time - start_time) / 60
            print(f"Dados atualizados com sucesso! Tempo de execução: {execution_time_minutes:.2f} minutos.")

        except Exception as e:
            self.conn.rollback()
            print(f"Erro durante a transação: {e}")

        finally:
            self.conn.close()

if __name__ == "__main__":
    from config import *
    
    atualizador = AtualizadorDadosCompras(dir_db_relatorio_tresmann, dir_db_preços_4_lojas)
    atualizador.atualizar_dados_compras()
