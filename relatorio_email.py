import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1

# ABRIR CHROME -
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")

# ABRIR SISTEMA
pyperclip.copy("https://drive.google.com/drive/folders/14oLE59U1RqyRqlBbKpsyymW-mitvbtoh")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(3)   

# FAZER DOWNLOAD
pyautogui.click(x=353, y=284)
pyautogui.click(x=1714, y=194)
pyautogui.click(x=1497, y=593)

# ANALISAR DADOS
dados = pd.read_excel(r"E:\DOWNLOADS\Vendas - Dez.xlsx")    
print(dados)
faturamento = dados['Valor Final'].sum()
quantidade = dados['Quantidade'].sum()
print(faturamento)
print(quantidade)

# ENVIAR E-MAIL
pyautogui.hotkey("ctrl", "t")
pyautogui.write("www.gmail.com")
pyautogui.press("enter")
time.sleep(5)
pyautogui.click(x=111, y=199)
pyperclip.copy("siteevandro@gmail.com")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
time.sleep(1)
pyautogui.write("Faturamento Dezembro")
pyautogui.press("tab")
texto = f"""
Prezados, bom dia

O Faturamento de dezembro foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Atenciosamente,
Evandro Perez
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")
pyautogui.alert("Automação Finalizada")