import smtplib
import ssl
import pandas as pd
import locale
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from config import email_remetente, email_senha, email_destinatarios

def excel_to_html(excel_path, sheet_name):
    # Carregar o arquivo Excel
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Substituir o valor NaN na quarta linha, primeira coluna por "Total:"
    df.iloc[3, 0] = 'Total:'

    # Definir a formatação de números no formato "R$1.000,00" e com separador de milhar
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    df['VALOR CUSTO TOTAL'] = df['VALOR CUSTO TOTAL'].apply(lambda x: locale.currency(x, grouping=True))
    df['VALOR VENDA TOTAL'] = df['VALOR VENDA TOTAL'].apply(lambda x: locale.currency(x, grouping=True))
    df['QTDE ITENS'] = df['QTDE ITENS'].apply(lambda x: f'{locale.format_string("%.0f", x, grouping=True)}')
    df['QTDE EXCEDENTE'] = df['QTDE EXCEDENTE'].apply(lambda x: f'{locale.format_string("%.0f", x, grouping=True)}')

    # Converter o DataFrame para HTML e adicionar classes CSS
    html_table = df.to_html(index=False, escape=False, classes='custom-table')

    # Estilizar a tabela com CSS
    estilo_css = """
    <style>
    .custom-table {
        font-family: Arial, sans-serif;
        border-collapse: collapse;
        width: 100%;
    }
    .custom-table th {
        background-color: #363636;
        color: white;
        text-align: left;
        padding: 8px;
        border-bottom: 2px solid #000;
    }
    .custom-table th, .custom-table td {
        border: 1px solid #dddddd;
        text-align: left;
        padding: 8px;
    }
    .custom-table tbody tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .custom-table td:first-child {
        font-weight: bold; /* Torna a coluna "Loja" em negrito */
    }
    </style>
    """

    return estilo_css + html_table

def enviar_email():
    # Caminho do arquivo Excel
    excel_file_path = r'F:\BI\ESTOQUE_EXCEDENTE.xlsx'
    sheet_name = 'RESUMO'

    # Converter a planilha em HTML com estilo personalizado
    corpo_do_email = excel_to_html(excel_file_path, sheet_name)

    # Dados do email
    mensagem = MIMEMultipart()
    mensagem["From"] = email_remetente
    mensagem["To"] = ", ".join(email_destinatarios)
    mensagem["Subject"] = "Relatório de Estoque Excedente"

    # Adicionar o HTML no corpo do email
    mensagem.attach(MIMEText(corpo_do_email, "html"))

    # Anexar o arquivo Excel
    with open(excel_file_path, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name="RELATORIO_EXTOQUE_EXCEDENTE.xlsx")
        part['Content-Disposition'] = 'attachment; filename="RELATORIO_EXTOQUE_EXCEDENTE.xlsx"'
        mensagem.attach(part)

    try:
        with smtplib.SMTP_SSL("mail.agoraa.com.br", 465, context=ssl.create_default_context()) as server:
            server.login(email_remetente, email_senha)
            server.sendmail(email_remetente, email_destinatarios, mensagem.as_string())
            print("Email Enviado!")
    except smtplib.SMTPException as e:
        print(f"Erro ao enviar email: {e}")

if __name__ == "__main__":
    # Chamar a função para enviar o email
    enviar_email()