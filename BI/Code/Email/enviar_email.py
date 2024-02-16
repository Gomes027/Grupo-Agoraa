import ssl
import locale
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from Email.config import CAMINHO_ARQUIVO_ESTOQUES, CAMINHO_ARQUIVO_DESCONTOS, CAMINHO_ARQUIVO_ENCARTES, CAMINHO_ARQUIVO_LOJAS, CAMINHO_ARQUIVO_VINHOS, CAMINHO_ARQUIVO_VOLUME, EMAIL_REMETENTE, EMAIL_SENHA, EMAIL_DESTINATARIOS

def excel_to_html(excel_path, sheet_name, columns_moeda, columns_str, aplicar_total=False):
    """
    Converte uma planilha Excel em uma tabela HTML formatada.

    :param excel_path: Caminho do arquivo Excel.
    :param sheet_name: Nome da aba a ser convertida.
    :param columns_moeda: Colunas para aplicar formatação de moeda.
    :param columns_str: Colunas para tratar como strings, aplicando agrupamento.
    :param aplicar_total: Se True, aplica a modificação 'Total:' na posição especificada.
    :return: String contendo o HTML da tabela.
    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

    for column in columns_moeda:
        df[column] = pd.to_numeric(df[column], errors='coerce').apply(lambda x: locale.currency(x, grouping=True) if pd.notnull(x) else "")

    for column in columns_str:
        df[column] = df[column].apply(lambda x: f'{locale.format_string("%.0f", x, grouping=True) if pd.notnull(x) else ""}' if isinstance(x, (int, float)) else x)

    if aplicar_total:
        # Aplicar tags <b> diretamente nos valores da última linha
        df.iloc[-1] = df.iloc[-1].apply(lambda x: f'<b>{x}</b>')

    html_table = df.to_html(index=False, escape=False, classes='custom-table', border=0)

    estilo_css = """
    <style>
        .custom-table {
            font-family: Arial, sans-serif;
            border-collapse: collapse;
            width: 100%;
        }
        .custom-table th, .custom-table td {
            border: 1px solid #dddddd;
            text-align: left;
            padding: 8px;
        }
        .custom-table th {
            background-color: #363636;
            color: white;
        }
        .custom-table tbody tr:nth-child(even) {
            background-color: #f2f2f2;
        }
    </style>
    """

    return estilo_css + html_table

def enviar_email():
    estoque_p1 = excel_to_html(CAMINHO_ARQUIVO_ESTOQUES, "RESUMO", ['VALOR CUSTO TOTAL', 'VALOR VENDA TOTAL'], ['QTDE ITENS', 'QTDE EXCEDENTE'], aplicar_total=True)
    descontos = excel_to_html(CAMINHO_ARQUIVO_DESCONTOS, "Sheet1", [], ["SMJ", "STT", "VIX", "Total Geral", "DESCONTOS"], aplicar_total=True)
    encartes = excel_to_html(CAMINHO_ARQUIVO_ENCARTES, "Resumo", [], ["INICIO", "FIM", "QTDE TOTAL", "VALOR TOTAL", "LUCRO BRUTO TOTAL", "% LUCRO", "LUCRO / PROD"], aplicar_total=True)
    lojas_controle = excel_to_html(CAMINHO_ARQUIVO_LOJAS, "LOJA CONTROLE", ["CUSTO"], ["LOJA", "QTDE REGISTROS", "SOMA PROD."], aplicar_total=True)
    lojas_trocas = excel_to_html(CAMINHO_ARQUIVO_LOJAS, "LOJA TROCAS", ["CUSTO"], ["LOJA", "QTDE REGISTROS", "SOMA PROD."], aplicar_total=True)
    vinhos = excel_to_html(CAMINHO_ARQUIVO_VINHOS, "RESUMO", [], ["DESCRIÇÃO", "SMJ", "STT", "VIX", "MCP", "TOTAL"], aplicar_total=False)
    volume_dos_100_smj = excel_to_html(CAMINHO_ARQUIVO_VOLUME, "TRESMANN - SMJ", [], ["PORCENTAGEM" , "QTDE ITENS", "VALOR DA VENDA",  "% / TOTAL"], aplicar_total=False)
    volume_dos_100_stt = excel_to_html(CAMINHO_ARQUIVO_VOLUME, "TRESMANN - STT", [], ["PORCENTAGEM" , "QTDE ITENS", "VALOR DA VENDA",  "% / TOTAL"], aplicar_total=False)
    volume_dos_100_vix = excel_to_html(CAMINHO_ARQUIVO_VOLUME, "TRESMANN - VIX", [], ["PORCENTAGEM" , "QTDE ITENS", "VALOR DA VENDA",  "% / TOTAL"], aplicar_total=False)

    # Adicionando títulos em negrito para cada tabela
    titulo_estoque = '<h2 style="font-weight:bold;">ESTOQUE EXCEDENTE</h2>'
    titulo_descontos = '<h2 style="font-weight:bold;">DESCONTOS NO PDV</h2>'
    titulo_encartes = '<h2 style="font-weight:bold;">ENCARTES</h2>'
    titulo_volumes = '<h2 style="font-weight:bold;">VOLUME DOS 100</h2>'
    titulo_controle = '<h2 style="font-weight:bold;">LOJAS DE CONTROLE</h2>'
    titulo_trocas = '<h2 style="font-weight:bold;">LOJAS DE TROCA</h2>'
    titulo_vinhos = '<h2 style="font-weight:bold;">VINHOS</h2>'

    # Construindo o corpo do email com as lojas TRESMANN em negrito
    corpo_email = (
        titulo_estoque + estoque_p1 + '<br><br><br>' +
        titulo_descontos + descontos + '<br><br><br>' +
        titulo_encartes + encartes + '<br><br><br>' +
        titulo_controle + lojas_controle + '<br>' +
        titulo_trocas + lojas_trocas + '<br><br><br>' +
        titulo_vinhos + vinhos + '<br><br><br>' +
        titulo_volumes +    '<strong>TRESMANN - SMJ</strong>' + volume_dos_100_smj + '<br>' +
                            '<strong>TRESMANN - STT</strong>' + volume_dos_100_stt + '<br>' +
                            '<strong>TRESMANN - VIX</strong>' + volume_dos_100_vix
    )
    
    mensagem = MIMEMultipart()
    mensagem["From"] = EMAIL_REMETENTE
    mensagem["To"] = ", ".join(EMAIL_DESTINATARIOS)
    mensagem["Subject"] = "RELATÓRIOS BI"

    mensagem.attach(MIMEText(corpo_email, "html"))

    try:
        with smtplib.SMTP_SSL("mail.agoraa.com.br", 465, context=ssl.create_default_context()) as server:
            server.login(EMAIL_REMETENTE, EMAIL_SENHA)
            server.sendmail(EMAIL_REMETENTE, EMAIL_DESTINATARIOS, mensagem.as_string())
            print("Email enviado com sucesso!")
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar email: {e}")