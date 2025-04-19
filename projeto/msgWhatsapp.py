import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui

#Abrir a planilha com os contatos
workbook = openpyxl.load_workbook(r'C:\Users\ruans\OneDrive\Documentos\Python Scripts\projeto\clientes.xlsx')
planilha = workbook['Sheet1']

#Percorrer cada linha da planilha
for line in planilha.iter_rows(min_row=2):
    name = line.value
    phone = line[1].value
    
    #Quando acabar os contatos ele encerra o script
    if name is None or phone is None:
        break
    
    #configuração de mensagem e do link da mensagem
    message = f'Hello {name}, This is an automated message'
    #codifica a mensagem para estar aceitável no link 
    link_mensagem = f'https://web.whatsapp.com/send?phone={phone}&text={quote(message)}'
    webbrowser.open(link_mensagem)
    sleep(35)
    pyautogui.click(x=1325, y=690)
    sleep(5)
    pyautogui.hotkey("ctrl","w")
    