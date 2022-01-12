import pyautogui 
import pyperclip
from time import sleep
import pandas as pd

pyautogui.PAUSE = 1.5
pyautogui.click(122, 877)

pyautogui.hotkey('ctrl', 't')
pyautogui.write('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga')
pyautogui.press('enter')
sleep(5)
pyautogui.click(331, 249, clicks=2)
sleep(2)
pyautogui.click(355, 253, button='right')
sleep(3)
pyautogui.click(491, 740)
sleep(5)
pyautogui.press("enter")
sleep(10)


tabela = pd.read_excel(r"C:\Users\galva\Downloads\Vendas - Dez.xlsx")

quantidade_produtos = tabela["Quantidade"].sum()
faturamento = tabela["Valor Final"].sum()

print(quantidade_produtos)
print(faturamento)

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.press("enter")
sleep(8)
pyautogui.click(79, 167)
pyperclip.copy("g.alvarez.b12@gmail.com")
sleep(2)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("Tab")
pyautogui.write("Intensivao de python Aula 1")
pyautogui.press("Tab")

texto = f'''
Prezado Sr Gustavo Alvarez

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade_produtos:,}

Eu msm
''' 
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
