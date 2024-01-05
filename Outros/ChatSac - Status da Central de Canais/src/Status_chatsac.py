import ssl
import smtplib
import schedule
from PIL import Image
from time import sleep
from datetime import date
from selenium import webdriver
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from config import *

def status_chatsac():
  # Opções
  chrome_options = Options() 
  chrome_options.add_argument("start-maximized") # Iniciar maximizado 

  # Inicia o WebDriver
  driver = webdriver.Chrome(options=chrome_options)
  driver.get(URL)
  driver.implicitly_wait(20)

  #Preenche os dados e faz login
  driver.find_element(By.XPATH, "//input[@type='text'][contains(@class, 'MuiInputBase-input')]").send_keys(EMAIL_LOGIN)
  driver.find_element(By.NAME, 'Senha').send_keys(SENHA_LOGIN)
  driver.find_element(By.XPATH, "//button[@class='MuiButtonBase-root MuiButton-root MuiButton-contained w-full MuiButton-containedPrimary MuiButton-containedSizeLarge MuiButton-sizeLarge']").click()
  driver.implicitly_wait(20)

  # Central de canais
  driver.find_element(By.XPATH, "//span[@class='cursor-pointer']").click()
  sleep(5)

  # Tira print e recorta a imagem
  driver.save_screenshot(DIR_IMAGEM)
  imagem = Image.open(DIR_IMAGEM)
  area_corte = (161, 76, 1120, 550)
  imagem_cortada = imagem.crop(area_corte)
  imagem_cortada.save(DIR_IMAGEM)

  driver.quit() # Fecha o webdriver

def enviar_email():
  # Dados do email
  mensagem = MIMEMultipart()
  mensagem["From"] = REMETENTE
  mensagem["To"] = ", ".join(DESTINATARIO)
  mensagem["Subject"] = f"ChatSac Status"

  with open(DIR_IMAGEM, "rb") as f:
      img_data = f.read()

  # Anexando a imagem para uso no HTML
  img_html = MIMEImage(img_data)
  img_html.add_header("Content-ID", "<CHATSAC>")
  mensagem.attach(img_html)

  # Anexando a imagem como um anexo regular
  img_anexo = MIMEImage(img_data)
  img_anexo.add_header("Content-Disposition", "attachment", filename="Chatsac.jpg")
  mensagem.attach(img_anexo)

  # Corpo do email em HTML
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

def job():
    print("Executando as tarefas agendadas...")
    try:
      status_chatsac()
    except:
      status_chatsac()

    enviar_email()

DIR_IMAGEM = r"Img\ChatSac.png"

# Agendar para executar todos os dias às 06:30 AM
schedule.every().day.at("06:00").do(job)

while True:
    schedule.run_pending()
    sleep(1)