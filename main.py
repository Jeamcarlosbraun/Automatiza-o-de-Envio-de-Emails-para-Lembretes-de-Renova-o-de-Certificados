# -*- coding: utf-8 -*-
import pyautogui
import time
import pandas as pd
import pyperclip
from datetime import datetime

try:
    tabela = pd.read_excel("Modelo.xlsx")
except FileNotFoundError:
    print("Arquivo Excel não encontrado. Verifique o caminho e o nome do arquivo.")
    exit()
    
def abrir_outlook():
    #Abre o Microsoft Outlook.
    pyautogui.press("win")
    pyautogui.write("Outlook")
    pyautogui.press("enter")
    time.sleep(3)
    
abrir_outlook()

# AItera sobre as linhas do arquivo Excel e preenche os campos do e-mail, extrai os dados
#e insere no email
for coluna in range(len(tabela)):
    contato = str(tabela["Contato"][coluna])
    nome = tabela["Nome"][coluna]
    data = tabela["Data"][coluna].strftime('%d/%m/%Y')
    tipo = tabela["Tipo"][coluna]
    if str(tipo) == "1":
        modelo = "(CNPJ)"
    else:
        modelo = "(ECPF)"
    pyautogui.click(x=19, y=109)
    pyautogui.write(contato)
    pyautogui.click(x=516, y=301)
    assunto = "RENOVAÇÃO CERTIFICADO DIGITAL"
    pyperclip.copy(assunto)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.5)
    pyautogui.press("tab")
    time.sleep(0.5)
    pyautogui.press("tab")
    escrita = "Bom Dia!\nConforme consta em nossos controles está vencendo o seguinte certificado digital:\n -  " + nome + " " + modelo + " " + data + "\nSugestão de contato da ############# para solicitar a renovação: #############.\nAssim que fizerem os trâmites se possível nos encaminhar uma cópia para mantermos em nossos controles.\nATT"
    print(escrita)
    pyperclip.copy(escrita)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.click(x=308, y=267)
    time.sleep(0.5)

#
pyautogui.alert("Emails enviados com sucesso!")