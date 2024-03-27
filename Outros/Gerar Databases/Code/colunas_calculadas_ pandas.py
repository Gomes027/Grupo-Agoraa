import sqlite3
import pandas as pd
import time

def get_precos(codigo, loja):
    precos_conn = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_precos_4_lojas.db')
    precos_cursor = precos_conn.cursor()
    precos_cursor.execute('''SELECT custo_liquido, preco_venda 
                             FROM precos_lojas 
                             WHERE codigo = ? AND nome_loja = ?''', (codigo, loja))
    row = precos_cursor.fetchone()
    precos_conn.close()
    return row

start_time = time.time()

# Conectar ao banco de dados
conn = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_tresmann.db')

# Consulta SQL para obter os dados
query = '''
    SELECT codigo, loja, estoque_atual, venda_ultimos_12_meses, venda_ultimos_30_dias, validade, qtde_embalagem
    FROM dados_compras
'''

# Ler os dados do banco de dados usando Pandas
df = pd.read_sql_query(query, conn)

# Calcular novas colunas
df['vda_media_30d_12m'] = df['venda_ultimos_12_meses'] / 12
df['maior_vda'] = df[['venda_ultimos_30_dias', 'vda_media_30d_12m']].max(axis=1)
df['venda_diaria'] = df['maior_vda'] / 30
df['validade'] = pd.to_datetime(df['validade'], format='%d/%m/%Y', errors='coerce')
df['validade'] = df['validade'].fillna(df['validade'].min())
df['dias_para_vencer'] = ((df['validade'] - pd.Timestamp.now()) / pd.Timedelta(days=1)).astype(int).clip(lower=0)
df['projecao_vendas'] = df['dias_para_vencer'] * df['venda_diaria']
df['est_proj'] = df['estoque_atual'].gt(df['projecao_vendas']).astype(int)
df['est_excedente'] = df['estoque_atual'].sub(df['projecao_vendas']).clip(lower=0)

# Fechar a conexão com o banco de dados
conn.close()

# Conectar novamente para atualizar o banco de dados
conn = sqlite3.connect(r'F:\COMPRAS\Automações.Compras\Bancos de Dados\relatorio_tresmann.db')

# Atualizar o banco de dados com os novos valores
df[['p_custo', 'p_venda']] = pd.DataFrame([get_precos(codigo, loja) for codigo, loja in df[['codigo', 'loja']].values], index=df.index)
df['total_custo'] = df['est_excedente'] * df['p_custo']
df['total_venda'] = df['est_excedente'] * df['p_venda']

# Atualizar o banco de dados
df[['p_custo', 'p_venda', 'total_custo', 'total_venda']].to_sql('dados_compras', conn, if_exists='replace', index=False)

# Fechar a conexão com o banco de dados
conn.close()

end_time = time.time()  # Captura o tempo de término
execution_time = end_time - start_time  # Calcula o tempo de execução
print(f"Dados atualizados com sucesso! Tempo de execução: {execution_time:.2f} segundos.")