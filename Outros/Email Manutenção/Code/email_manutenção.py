import os
import smtplib
import ssl
from email.message import EmailMessage
import mimetypes
from config import *

class EmailComAnexos:
    def __init__(self, servidor, porta, email_remetente, email_senha, pasta_de_arquivos, email_destinatarios):
        self.servidor = servidor
        self.porta = porta
        self.email_remetente = email_remetente
        self.email_senha = email_senha
        self.pasta_de_arquivos = pasta_de_arquivos
        self.email_destinatarios = email_destinatarios

    def criar_mensagem_com_anexos(self, assunto, corpo):
        mensagem = EmailMessage()
        mensagem["From"] = self.email_remetente
        mensagem["To"] = self.email_destinatarios
        mensagem["Subject"] = assunto
        mensagem.set_content(corpo)

        # Anexar arquivos .xlsx da pasta especificada
        for arquivo in os.listdir(self.pasta_de_arquivos):
            caminho_completo = os.path.join(self.pasta_de_arquivos, arquivo)
            if os.path.isfile(caminho_completo) and arquivo.endswith(".xlsx"):
                tipo_mime, _ = mimetypes.guess_type(arquivo)
                tipo_mime_principal, subtipo_mime = tipo_mime.split('/', 1)

                with open(caminho_completo, 'rb') as arquivo_aberto:
                    mensagem.add_attachment(arquivo_aberto.read(),
                                            maintype=tipo_mime_principal,
                                            subtype=subtipo_mime,
                                            filename=arquivo)

        return mensagem

    def enviar_email(self, assunto, corpo):
        mensagem = self.criar_mensagem_com_anexos(assunto, corpo)
        context = ssl.create_default_context()

        try:
            with smtplib.SMTP_SSL(self.servidor, self.porta, context=context) as server:
                server.login(self.email_remetente, self.email_senha)
                server.send_message(mensagem)
                print("Email Enviado!")
        except smtplib.SMTPException as e:
            print(f"Erro ao enviar email: {e}")

if __name__ == "__main__":
    email_enviador = EmailComAnexos(servidor, porta, email_remetente, email_senha, pasta_de_arquivos, email_destinatarios)
    email_enviador.enviar_email("Relatório atualizado de controle de acessos e equipamentos", "Segue relátorios em anexo:")