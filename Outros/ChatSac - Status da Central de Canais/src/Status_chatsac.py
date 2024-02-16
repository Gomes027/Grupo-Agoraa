import asyncio
import smtplib
import ssl
import schedule
from time import sleep
from datetime import date
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from pyppeteer import launch
from pyppeteer.errors import TimeoutError
from config import EMAIL_LOGIN, SENHA_LOGIN, URL, DIR_IMAGEM, REMETENTE, DESTINATARIO, SENHA_REMETENTE

async def status_chatsac():
  browser = None
  try:
      browser = await launch(headless=False, args=['--start-maximized'])
      page = await browser.newPage()
      # Defina a viewport para a resolução desejada, por exemplo, 1920x1080
      await page.setViewport({'width': 1280, 'height': 960})
      await page.goto(URL)
      await page.waitForSelector("input[type='text']")
      await page.type("input[type='text']", EMAIL_LOGIN)
      await page.type("input[name='Senha']", SENHA_LOGIN)
      login_button_selector = "button[class='MuiButtonBase-root MuiButton-root MuiButton-contained w-full MuiButton-containedPrimary MuiButton-containedSizeLarge MuiButton-sizeLarge']"
      await page.click(login_button_selector)

      # Clique no elemento SVG específico
      svg_selector = "svg[data-icon='plug'].svg-inline--fa.fa-plug.fa-lg"
      await page.waitForSelector(svg_selector)
      await page.click(svg_selector)

      popup_selector = ".MuiDialog-paper"
      await page.waitForSelector(popup_selector, {'visible': True, 'timeout': 60000})
      await asyncio.sleep(5)
      element = await page.querySelector(popup_selector)
      await element.screenshot({'path': DIR_IMAGEM})
      return True
  except Exception as e:
      print(f"Erro ao executar status_chatsac: {e}")
      return False
  finally:
      if browser is not None:
          await browser.close()

def enviar_email():
  mensagem = MIMEMultipart()
  mensagem["From"] = REMETENTE
  mensagem["To"] = ", ".join(DESTINATARIO)
  mensagem["Subject"] = "ChatSac Status"
  
  with open(DIR_IMAGEM, "rb") as f:
      img_data = f.read()
  
  img_html = MIMEImage(img_data)
  img_html.add_header("Content-ID", "<CHATSAC>")
  mensagem.attach(img_html)
  
  img_anexo = MIMEImage(img_data)
  img_anexo.add_header("Content-Disposition", "attachment", filename="Chatsac.jpg")
  mensagem.attach(img_anexo)
  
  html = """\
  <html>
    <body>
      <p>Central de Canais do ChatSac:</p>
      <img src="cid:CHATSAC">
    </body>
  </html>
  """
  parte_html = MIMEText(html, "html")
  mensagem.attach(parte_html)
  
  try:
      with smtplib.SMTP_SSL("mail.agoraa.com.br", 465, context=ssl.create_default_context()) as server:
          server.login(REMETENTE, SENHA_REMETENTE)
          server.sendmail(REMETENTE, DESTINATARIO, mensagem.as_string())
  except smtplib.SMTPException as e:
      print(f"Erro ao enviar email: {e}")

def main():
    download_sucesso = False

    while not download_sucesso:
        download_sucesso = asyncio.run(status_chatsac())

    enviar_email()

if __name__ == '__main__':
    main()