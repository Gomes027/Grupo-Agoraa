import sqlite3
import pandas as pd
import time

start_time = time.time()  # Captura o tempo de início

# Conectar ao banco de dados
conn = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_tresmann.db')
cursor = conn.cursor()

# Adicionar colunas
"""cursor.execute('''ALTER TABLE dados_compras ADD COLUMN vda_media_30d_12m REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN maior_vda REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN venda_diaria REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN dias_para_vencer INTEGER''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN projecao_vendas REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN est_proj INTEGER''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN est_excedente REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN p_custo REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN p_venda REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN total_custo REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN total_venda REAL''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN prazo TEXT''')
cursor.execute('''ALTER TABLE dados_compras ADD COLUMN est_atual_x_emb TEXT''')"""

# Calcular e atualizar os valores das novas colunas
cursor.execute('''UPDATE dados_compras 
                  SET vda_media_30d_12m = venda_ultimos_12_meses / 12,
                      maior_vda = CASE WHEN venda_ultimos_30_dias > (venda_ultimos_12_meses / 12) THEN venda_ultimos_30_dias ELSE (venda_ultimos_12_meses / 12) END''')

cursor.execute('''UPDATE dados_compras 
                  SET venda_diaria = maior_vda / 30,
                      dias_para_vencer = CASE
                            WHEN (strftime('%s', validade) - strftime('%s', 'now', '-1 day')) / 86400 < 0 THEN 0
                            ELSE CAST((strftime('%s', validade) - strftime('%s', 'now', '-1 day')) / 86400 AS INTEGER)
                            END -- diferença em dias, garantindo que não seja negativo''')

cursor.execute('''UPDATE dados_compras 
                  SET projecao_vendas = dias_para_vencer * venda_diaria''')

cursor.execute('''UPDATE dados_compras                   
                    SET est_proj = CASE WHEN estoque_atual > projecao_vendas THEN 1 ELSE 0 END,
                        est_excedente = CASE WHEN est_proj = 0 THEN 0 ELSE MAX(estoque_atual - projecao_vendas, 0) END''')

# Função para obter P. CUSTO e P.VENDA
def get_precos(codigo, loja):
    precos_conn = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_precos_4_lojas.db')
    precos_cursor = precos_conn.cursor()
    precos_cursor.execute('''SELECT custo_liquido, preco_venda 
                             FROM precos_lojas 
                             WHERE codigo = ? AND nome_loja = ?''', (codigo, loja))
    row = precos_cursor.fetchone()
    precos_conn.close()
    return row

# Dicionário para armazenar os preços
precos_dict = {}

# Coletar os preços
for row in cursor.execute('SELECT codigo, loja FROM dados_compras'):
    codigo, loja = row
    print(f"Código: {codigo}, Loja: {loja}")
    
    precos = get_precos(codigo, loja)
    if precos:
        p_custo, p_venda = precos
        precos_dict[(codigo, loja)] = (p_custo, p_venda)
        print(f"Preços obtidos: {precos}")  # Verifica os valores obtidos da função get_precos
        print(f"P. Custo: {p_custo}, P. Venda: {p_venda}")  # Verifica os valores a serem atualizados
    else:
        print("Nenhum preço encontrado para este código e loja.")

try:
    # Atualizar os registros em lotes
    for (codigo, loja), (p_custo, p_venda) in precos_dict.items():
        cursor.execute('''UPDATE dados_compras 
                          SET p_custo = ?,
                              p_venda = ?
                          WHERE codigo = ? AND loja = ?''', (p_custo, p_venda, codigo, loja))

    # Calcular TOTAL CUSTO e TOTAL VENDA
    cursor.execute('''UPDATE dados_compras 
                      SET total_custo = est_excedente * p_custo,
                          total_venda = est_excedente * p_venda''')
    
    # Atualizar PRAZO e EST ATUAL < X EMB
    cursor.execute('''UPDATE dados_compras 
                    SET prazo = CASE WHEN dias_para_vencer < 60 THEN 'SIM' ELSE 'NÃO' END,
                        est_atual_x_emb = CASE WHEN estoque_atual < qtde_embalagem THEN 'SIM' ELSE 'NÃO' END''')
    
    print("Dados atualizados com sucesso!")
    
    conn.commit()  # Confirmar transação se tudo ocorrer bem

except Exception as e:
    conn.rollback()  # Reverter transação em caso de erro
    print(f"Erro durante a transação: {e}")

finally:
    conn.close()  # Fechar a conexão com o banco de dados

end_time = time.time()  # Captura o tempo de término
execution_time = end_time - start_time  # Calcula o tempo de execução
print(f"Dados atualizados com sucesso! Tempo de execução: {execution_time:.2f} segundos.")