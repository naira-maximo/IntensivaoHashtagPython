# Automação de Sistemas e Processos com Python

# Desafio:

# Todos os dias, o sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing

# Para resolver isso, usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

import pyautogui
import pyperclip
import time

pyautogui.PAUSE = 1

# pyautogui.click -> clicar
# pyautogui.press -> apertar apenas uma tecla
# pyautogui.hotkey -> apertar um conjunto de teclas (atalho)
# pyautogui.write -> escrever um texto

# Caso o navegador não esteja aberto, executar essa sequência de comandos:
pyautogui.press('win')
pyautogui.write('opera') # Digitar o nome do navegador
pyautogui.press('enter')
time.sleep(2) # Esperar um tempo para o navegador abrir

# Passo 1: Entrar no sistema da empresa (como exemplo, link do Drive)
# pyautogui.hotkey('ctrl', 't') # Abrir uma nova aba no navegador
pyperclip.copy('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')

time.sleep(5) # Tempo para a tela carregar

# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)

# Código para descobrir qual a posição de um item que queira clicar
# pyautogui.click(x=302, y=343, clicks=2) # x e y são os parâmetros para a localização do cursos na tela.
# clicks é o parâmetro para a quantidade de clicks e button é o parâmetro para qual botão do mouse clicar (left, right e center)
# time.sleep(5) # Tempo para posicionar o cursor

# Passo 3: Fazer o download da base de vendas
pyautogui.click(x=611, y=437, clicks=2) # clicar para abrir a pasta "Executar"
time.sleep(2) # Esperar um tempo para a pasta abrir
pyautogui.click(x=549, y=528) # clicar no arquivo de Vendas
pyautogui.click(x=1677, y=202)
pyautogui.click(x=1359, y=707) # clicar para fazer download no arquivo
time.sleep(5) # esperar o download acabar

# Ler o arquivo baixado para pegar os indicadores
# - Faturamento
# - Quantidade de Produtos

# Passo 4: Importar a base de vendas pro Python
import pandas as pd

tabela = pd.read_excel(r"C:\Users\naira\Downloads\Vendas - Dez.xlsx")
print(tabela)

faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()

# Enviar um e-mail pelo gmail

# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda

# abrir aba
pyautogui.hotkey("ctrl", "t")

# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever
pyautogui.click(x=129, y=224)

time.sleep(2) # Tempo para a tela carregar

# preencher as informações do e-mail
pyautogui.write("seugmail@gmail.com")
pyautogui.press("tab") # selecionar o email

pyautogui.press("tab") # pular para o campo de assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") # pular para o campo de corpo do email

texto = f"""
Prezados,

Segue relatório de vendas:
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Me coloco à disposição.
Atenciosamente,
Analista de Dados
"""

# formatação dos números (moeda, dinheiro)

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")