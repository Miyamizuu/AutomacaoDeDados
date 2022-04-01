
#Olá, aqui está uma breve explicação de automação de sistemas e processos.

import time

# Entar no sistema da empresa (no nosso caso o drive)
import pyautogui
import pyperclip

pyautogui.PAUSE = 1

print(pyautogui.press("win"))
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

#Navegar no sistema até encontrar a base de dados
pyautogui.click(x=434, y=304, clicks=2)
pyautogui.sleep(2)

#Exportar a base de dados
pyautogui.click(x=366, y=436)
pyautogui.click(x=1716, y=205)
pyautogui.click(x=1552, y=591)
time.sleep(2)

#Calcular os indicadores
import pandas as pd

tabela = pd.read_excel(r"D:\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

#automatizar o envio de email
import pyautogui
import pyperclip

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=21, y=195)
time.sleep(3)
pyautogui.write("vitoriacanon1212@gmail.com")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
texto = f"""
Prezados, bom dia 

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,.2f}

Abraços
VitóriaPython
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")

