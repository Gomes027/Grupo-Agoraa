import os
import ssl
import smtplib
import pandas as pd
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from config import email_remetente, email_senha, email_destinatarios

def analise_relatorio_tresmann(ARQUIVO):
    # Carregar o DataFrame
    df = pd.read_excel(ARQUIVO)

    # Mapeamento dos nomes das lojas
    mapeamento_lojas = {
        "MERCAPP": "MCP",
        "TRESMANN - VIX": "VIX",
        "TRESMANN - SMJ": "SMJ",
        "TRESMANN - STT": "STT"
    }

    # Atualizar os nomes das lojas no DataFrame
    df['LOJA'] = df['LOJA'].replace(mapeamento_lojas)

    # Qtde Linhas totais
    total_linhas = len(df)
    total_linhas = "{:,}".format(total_linhas).replace(",", ".")

    # Qtde Linhas por loja
    linhas_por_loja = df["LOJA"].value_counts().to_dict()
    linhas_por_loja = {loja: "{:,}".format(linhas).replace(",", ".") for loja, linhas in linhas_por_loja.items()}
    linhas_por_loja = "<br>" + "<br>".join([f"{loja}: {linhas}" for loja, linhas in linhas_por_loja.items()])

    # Agrupar por código e nome do produto e contar as ocorrências
    contagem_produtos = df.groupby(['CODIGO', 'NOME']).size()

    # Filtrar grupos que não aparecem exatamente quatro vezes
    produtos_nao_repetidos_4x = contagem_produtos[contagem_produtos != 4]

    # Qtde Colunas
    total_colunas = df.shape[1]

    # Última linha com codigo e nome do produto, desconsiderando hortifruti e material de consumo.
    df_filtrado = df[~df['SETOR'].isin(['HORTIFRUTI', 'MATERIAL CONSUMO'])]
    descrição_ultima_linha = df_filtrado.iloc[-1, :2].to_dict()
    descrição_ultima_linha = '\n' + ' - '.join([str(valor) for valor in descrição_ultima_linha.values()])

    # Criando o dicionário de resposta
    resposta = {
        "Quantidade de Colunas": total_colunas,
        "Total de Linhas": total_linhas,
        "Linhas por Loja": linhas_por_loja,
        "Último Produto": descrição_ultima_linha,
    }

    # Verifica se há produtos que não se repetem exatamente quatro vezes
    if not produtos_nao_repetidos_4x.empty:
        # Formatando os resultados no formato "codigo - nome do produto"
        produtos_nao_repetidos_4x_formatados = [' - '.join(map(str, index)) for index in produtos_nao_repetidos_4x.index]
        produtos_nao_repetidos_4x_str = '<br>' + '<br>'.join(produtos_nao_repetidos_4x_formatados)

        resposta["Produtos com erro de Repetição"] = produtos_nao_repetidos_4x_str

    return resposta

def analise_historico_pedidos(ARQUIVO):
    # Carregar o DataFrame
    df_historico_pedidos = pd.read_excel(ARQUIVO)

    # Mapeamento dos nomes das lojas
    mapeamento_lojas = {
        "MERCAPP": "MCP",
        "TRESMANN - VIX": "VIX",
        "TRESMANN - SMJ": "SMJ",
        "TRESMANN - STT": "STT"
    }

    # Atualizar os nomes das lojas no DataFrame
    df_historico_pedidos['Nome da Loja'] = df_historico_pedidos['Nome da Loja'].replace(mapeamento_lojas)

    """ Quantidade total de pedidos ontem. """
    # Filtrar o DataFrame para incluir apenas as linhas onde 'Nome do Comprador' é diferente de 'LUCAS'
    df_historico_pedidos = df_historico_pedidos[df_historico_pedidos['Nome do Comprador'] != 'LUCAS']

    # Garantir que a coluna 'Data' esteja no formato de data correto
    df_historico_pedidos['Data'] = pd.to_datetime(df_historico_pedidos['Data'])

    # Ordenar o DataFrame pela coluna 'Pedido' em ordem decrescente
    df_ordenado = df_historico_pedidos.sort_values(by='Pedido', ascending=False)

    # Obter a primeira linha do DataFrame ordenado
    primeira_linha = df_ordenado.iloc[0]

    # Extrair os valores desejados
    pedido_maior = primeira_linha['Pedido']
    fornecedor = primeira_linha['Fornecedor']
    nome_da_loja = primeira_linha['Nome da Loja']

    """ Quantidade total de pedidos no mês. """
    # Identificar a data mais recente no DataFrame
    data_mais_recente = df_historico_pedidos['Data'].max().date()

    # Filtrar o DataFrame pela data mais recente
    pedidos_data_recente = df_historico_pedidos[df_historico_pedidos['Data'].dt.date == data_mais_recente]

    # Contar o número de pedidos na data mais recente
    quantidade_pedidos_data_recente = len(pedidos_data_recente)

    # Garantir que a coluna 'Data' esteja no formato de data correto
    df_historico_pedidos['Data'] = pd.to_datetime(df_historico_pedidos['Data'])

    # Identificar o mês e ano atuais
    mes_atual = datetime.now().month
    ano_atual = datetime.now().year
    
    # Filtrar o DataFrame pelo mês e ano atuais
    pedidos_mes_atual = df_historico_pedidos[(df_historico_pedidos['Data'].dt.month == mes_atual) & 
                                            (df_historico_pedidos['Data'].dt.year == ano_atual)]

    # Contar o número total de pedidos no mês atual
    quantidade_pedidos_mes_atual = len(pedidos_mes_atual)

    """ Quantidade total de pedidos. """
    # Contar o número total de pedidos
    quantidade_total_pedidos = len(df_historico_pedidos)

    ultimo_pedido = f"{pedido_maior} - {fornecedor} - {nome_da_loja}"

    return {
        f"Quantidade de pedidos em {data_mais_recente}": quantidade_pedidos_data_recente,
        "Quantidade de pedidos no mês atual": quantidade_pedidos_mes_atual,
        "Quantidade total de pedidos": quantidade_total_pedidos,
        "Último Pedido": ultimo_pedido,
    }

def listar_arquivos(diretorio, arquivo_adicional=None):
    informacoes_arquivos = []

    # Função para obter informações do arquivo
    def obter_informacoes_arquivo(caminho_arquivo):
        nome_arquivo, extensao = os.path.splitext(os.path.basename(caminho_arquivo))
        ultima_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho_arquivo))
        return {"nome": nome_arquivo, "data": ultima_modificacao}

    # Adiciona o arquivo adicional, se fornecido
    if arquivo_adicional and os.path.isfile(arquivo_adicional):
        informacoes_arquivos.append(obter_informacoes_arquivo(arquivo_adicional))

    # Adiciona os arquivos do diretório
    for arquivo in os.listdir(diretorio):
        caminho_completo = os.path.join(diretorio, arquivo)
        if os.path.isfile(caminho_completo):
            informacoes_arquivos.append(obter_informacoes_arquivo(caminho_completo))

    # Ordena os arquivos por data, do mais novo para o mais antigo
    informacoes_arquivos.sort(key=lambda x: x["data"], reverse=True)

    # Formata as informações para exibição
    informacoes_formatadas = [f"{info['data'].strftime('%d/%m/%Y %H:%M')} - {info['nome']}" for info in informacoes_arquivos]
    return "<br>".join(informacoes_formatadas)

def enviar_email():
    print("Executando a análise de dados e enviando e-mail...")

    # Chama as funções que leem os arquivos e retornam análises
    analise_tresmann = analise_relatorio_tresmann(r"F:\COMPRAS\relatorio_tresmann.xlsx")
    analise_pedidos = analise_historico_pedidos(r"F:\COMPRAS\historico de pedidos _oficial.xlsx")
    lista_arquivos = listar_arquivos(r"F:\BI\Bases", r"F:\COMPRAS\Analise compra por estoque v4.11.xlsb")

    print("Dados analisados")

    # Convertendo os dicionários em string para o corpo do email
    corpo_email_tresmann = "<br>".join([f"{k}: {v}" for k, v in analise_tresmann.items()])
    corpo_email_pedidos = "<br>".join([f"{k}: {v}" for k, v in analise_pedidos.items()])

    # Combinando os resultados em um único corpo de e-mail
    corpo_do_email = f"<b>Última Modificação das Bases:</b><br>{lista_arquivos}<br><br><b>Análise Relatório Tresmann:</b><br>{corpo_email_tresmann}<br><br><b>Análise Histórico de Pedidos:</b><br>{corpo_email_pedidos}"

    # Dados do email
    mensagem = MIMEMultipart()
    mensagem["From"] = email_remetente
    mensagem["To"] = ", ".join(email_destinatarios)
    mensagem["Subject"] = "Relatório de Validações Iniciais"
    mensagem.attach(MIMEText(corpo_do_email, "html"))

    try:
        with smtplib.SMTP_SSL("mail.agoraa.com.br", 465, context=ssl.create_default_context()) as server:
            server.login(email_remetente, email_senha)
            server.sendmail(email_remetente, email_destinatarios, mensagem.as_string())
            print("Email Enviado!")
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar email: {e}")

if __name__ == "__main__":
    enviar_email()