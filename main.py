# Importa as bibioteclas que vamos utilizar no código
from IPython.display import display
import pandas as pd
import pyautogui
import pyperclip
import time

# Assume uma pausa de 1 segundo entre cada ação
#pyautogui.PAUSE = 1

# Dá um alerta que o programa vai ser iniciado e informa para abrir o chrome
#pyautogui.alert('O programa de automatização vai ser iniciado feche o chrome e não mexa mais no teclado nem mouse')

# Abre o chrome
#pyautogui.click(510, 1022, clicks=2)

# Abre uma nova guia
#pyautogui.hotkey('ctrl', 't')

# Acessa a tabela para fazer o download
#linkExcel = 'https://drive.google.com/drive/u/0/folders/1FjQym-r3xwBb-ggOReQXux4GY7ALbXd8'
#pyperclip.copy(linkExcel)
#pyautogui.hotkey('ctrl', 'v')
#pyautogui.hotkey('enter')
#time.sleep(3)
#pyautogui.click(386, 405)
#pyautogui.click(1468, 147)
#pyautogui.click(1303, 526)
time.sleep(7)

# Utilizamos o modulo pandas para ler a tabela
df = pd.read_excel(r'C:\Users\Usuario\Downloads\TESTE PROG PYTHON TABELA.xlsx')

# Aqui pegamos o numero de linhas do excel para fazer uma repetição com linhas como condição
linhas = int(df.shape[0])
linhas = linhas - 1
while linhas>-1:
    #Ejeta as variaveis
    preco = 0.74
    email = df.at[linhas,'Endereço de e-mail']
    horasl = df.at[linhas,'Quantas horas mais ou menos você deixa ligada sua luz?']

    #Contas do valor com diversos watts
    w6 = 6*horasl*30/1000*preco
    w8 = 8*horasl*30/1000*preco
    w10 = 10*horasl*30/1000*preco
    w12 = 12*horasl*30/1000*preco
    w15 = 15*horasl*30/1000*preco
    w17 = 17*horasl*30/1000*preco
    w20 = 20*horasl*30/1000*preco
    w25 = 25*horasl*30/1000*preco
    w60 = 60*horasl*30/1000*preco
    w80 = 80*horasl*30/1000*preco
    w100 = 100*horasl*30/1000*preco
    w120 = 120*horasl*30/1000*preco
    w150 = 150*horasl*30/1000*preco

    print("Com uma lâmpada de 6 watts você paga por mês {0:.2f}".format(w8), w8)

    pyautogui.hotkey('ctrl', 't')

    linhas = linhas-1



