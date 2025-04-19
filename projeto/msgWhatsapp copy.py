import openpyxl
from urllib.parse import quote
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

options = Options()

options.add_argument(r"user-data-dir=C:\Users\ruans\AppData\Local\Google\Chrome\User Data")

options.add_argument("profile-directory=Profile 8")

options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--disable-dev-shm-usage")

#Abrir a planilha com os contatos
workbook = openpyxl.load_workbook(r'C:\Users\ruans\OneDrive\Documentos\Python Scripts\projeto\clientes.xlsx')
planilha = workbook['Sheet1']

#Percorrer cada linha da planilha
for linha in planilha.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    
    #Quando acabar os contatos ele encerra o script
    if nome is None or telefone is None:
        break
    
    #configuração de mensagem e do link da mensagem
    mensagem = f'Olá {nome}, essa é uma mensagem automática'
    #codifica a mensagem para estar aceitável no link 
    link_mensagem = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
    driver = webdriver.Chrome(options=options)
    driver.get(link_mensagem)

    