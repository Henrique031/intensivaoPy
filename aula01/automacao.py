import pyautogui
import pyperclip
import time
import pandas as pd

#pip install opencv-pythom # para o pythom ler alguma print na tela


# pyautogui.press('nome da tecla') pressiona uma telcla
# pyautogui.click(x = valorx y = valory) clica com o mause, no canto de sua preferencia
# pyautogui.write() escreve na tela
# pyautogui.hotkey() atalhos no teclado
# pyautogui.position() 

# pyperclip.copy("Esse texto entre parentes será copiado") copia caracteres

# time.sleep(5) tempo que ficara dormindo

# print(pyautogui.position()) #Comando para saber o posição do mouse

# pyautogui.PAUSE = 2

pyautogui.press("win")
time.sleep(2)
pyautogui.write("Microsoft Edge")
pyautogui.press("enter")

#Entrar no drive
# contador = 0
# while not pyautogui.locateOnScreen("drive.png", confidence=0.8):
#     contador += 1
#     time.sleep(1)
#     print(contador)
pyautogui.click(x=891, y=750)
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
pyautogui.click(x=298, y=267, clicks=2 ) # extrato
time.sleep(4)
pyautogui.click(x=290, y=334) #click panilha
pyautogui.click(x=1155, y=174) # 3 pontos
pyautogui.click(x=970, y=496) # download
time.sleep(6)
pyautogui.click(x=1078, y=152) # salvar download
time.sleep(5)
pyautogui.click(x=79, y=392) # escoha src
pyautogui.click(x=1194, y=696) # salvar

# Processamento da tabela exel
tabela = pd.read_excel(r'D:\rique\Downloads\Vendas - Dez.xlsx')
print(tabela)

# .sum() soma
# .count() contar
# .mean() média
faturamento = tabela['Valor Final'].sum()
qtde = tabela['Quantidade'].sum()

# Abrir o Gmail
pyautogui.click(x=891, y=750)
pyperclip.copy("https://mail.google.com/mail/u/0/?hl=pt-BR#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(8)

#Enviar gmail
pyautogui.click(x=91, y=198) #Clicar no botão Escrever
time.sleep(7)
pyautogui.write('antoniafernandesdeoliveira712@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey("ctrl", "v")
pyautogui.press('tab')

txt = f'''Prezados,

Segue o relátorio de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {qtde:,}

Qualquer dúvida, estou a disposicao.
Att.,
Henrique aluno.
'''

pyperclip.copy(txt)
pyautogui.hotkey("ctrl", "v")

#botão enviar
# pyautogui.hotkey("ctrl", "enter")




