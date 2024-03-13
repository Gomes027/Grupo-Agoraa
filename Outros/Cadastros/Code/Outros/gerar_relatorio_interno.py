import pandas as pd
from sqlalchemy import create_engine

def processar_relatorio(caminho_entrada, nome_tabela, db_path='sqlite:///DataBase/dados_internos.db'):
    """
    Processa o relatório de um arquivo Excel, mantendo apenas as colunas "CODG DE BARRAS",
    "CODIGO" e "NOME", convertendo "CODG DE BARRAS" e "CODIGO" para inteiros, removendo
    duplicatas e linhas com "CODG DE BARRAS" nulo ou 0, e salvando o resultado em um
    banco de dados SQLite.

    Parâmetros:
    caminho_entrada (str): O caminho do arquivo Excel de entrada.
    nome_tabela (str): O nome da tabela no banco de dados onde o resultado será salvo.
    db_path (str): Caminho para o arquivo do banco de dados SQLite (usando a URI do SQLAlchemy).
    """
    # Carregar a planilha do Excel
    df = pd.read_excel(caminho_entrada, sheet_name="relatorios_tresmann")

    df = df[(df["SETOR"] != "HORTIFRUTI") & (df["SETOR"] != "MATERIAL CONSUMO")]

    # Manter apenas as colunas desejadas
    df = df[["CODG DE BARRAS", "CODIGO", "NOME"]]

    # Converter "CODG DE BARRAS" e "CODIGO" para inteiros
    df["CODG DE BARRAS"] = pd.to_numeric(df["CODG DE BARRAS"], errors='coerce').fillna(0).astype('int64')
    df["CODIGO"] = pd.to_numeric(df["CODIGO"], errors='coerce').fillna(0).astype(int)

    # Remover duplicatas com base na coluna "CODIGO"
    df = df.drop_duplicates(subset=["CODIGO"])

    # Filtrar linhas onde "CODG DE BARRAS" não é nulo ou 0
    df = df[(df["CODG DE BARRAS"].notnull()) & (df["CODG DE BARRAS"] != 0)]

    # Remover linhas em branco
    df = df.dropna(how='all')

    # Ordenar linhas pela coluna "CODG DE BARRAS"
    df = df.sort_values(by=["CODG DE BARRAS"])

    # Criar conexão ao banco de dados SQLite
    engine = create_engine(db_path)
    df.to_sql(name=nome_tabela, con=engine, index=False, if_exists='replace')

def consultar_dados(db_path='sqlite:///DataBase/dados_internos.db', nome_tabela='codigos_internos'):
    """
    Consulta dados de uma tabela específica em um banco de dados SQLite e retorna um DataFrame pandas com os resultados.

    Parâmetros:
    db_path (str): Caminho para o arquivo do banco de dados SQLite (usando a URI do SQLAlchemy).
    nome_tabela (str): O nome da tabela no banco de dados de onde os dados serão consultados.
    """
    
    # Criar conexão ao banco de dados SQLite
    engine = create_engine(db_path)
    
    # Consultar todos os dados da tabela especificada e carregar em um DataFrame
    df = pd.read_sql_table(nome_tabela, con=engine)
    
    return df

if __name__ == "__main__":
    processar_relatorio(r"F:\COMPRAS\relatorio_tresmann.xlsx", "codigos_internos")
    df_resultado = consultar_dados()
    print(df_resultado)
    