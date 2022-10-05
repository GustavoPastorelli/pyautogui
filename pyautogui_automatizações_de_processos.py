#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Todos os dias, o nosso sistema atualiza as vendas do dia anterior.
# O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior
# 
# E-mail da diretoria: seugmail+diretoria@gmail.com<br>
# Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
# 
# Referência do pyautogui: https://pyautogui.readthedocs.io/en/latest/quickstart.html

# In[33]:


get_ipython().system('pip install pyautogui')


# In[34]:


import pyautogui
import pyperclip
import time


# In[46]:



pyautogui.PAUSE = 1

# pyautogui.click -> clicar
# pyautogui.press -> apertar 1 tecla
# pyautogui.hotkey -> conjunto de teclas
# pyautogui.write -> escreve um texto

# Passo 1: Entrar no sistema da empresa (no nosso caso é o link do drive)
pyautogui.hotkey("ctrl", "t")
pyautogui.click(x=719, y=52)
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")



time.sleep(5)

# Passo 2: Navegar no sistema e encontrar a base de vendas (entrar na pasta exportar)
pyautogui.click(x=332, y=260, clicks=2)
time.sleep(2)

pyautogui.click(x=332, y=260, clicks=1)

#3 pontinhos
pyautogui.click(x=1389, y=163)

#baixando arquivo
pyautogui.click(x=1231, y=553)

time.sleep(5) # esperar o download acabar


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[47]:


# Passo 4: Importar a base de vendas pro Python
import pandas as pd

tabela = pd.read_excel(r"C:\Users\gustavo.aguiar\Downloads\Vendas - Dez.xlsx")
display(tabela)


# In[48]:


# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
print(faturamento)
quantidade = tabela["Quantidade"].sum()
print(quantidade)


# ### Vamos agora enviar um e-mail pelo gmail

# In[53]:


# Passo 6: Enviar um e-mail para a diretoria com os indicadores de venda

# abrir aba
pyautogui.hotkey("ctrl", "t")
pyautogui.click(x=719, y=52)
# entrar no link do email - https://mail.google.com/mail/u/0/#inbox
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever
pyautogui.click(x=101, y=167)

# preencher as informações do e-mail
pyautogui.write("pythonimpressionador@gmail.com")
time.sleep(3)
pyautogui.press("tab") # selecionar o email
time.sleep(3)

time.sleep(3)
pyperclip.copy("Relatório de Vendas")
time.sleep(3)
pyautogui.hotkey("ctrl", "v")
time.sleep(3)

pyautogui.press("tab") # pular para o campo de corpo do email
time.sleep(3)

texto = f"""
Prezados,

Segue relatório de vendas.
Faturamento: R${faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida estou à disposição.
Att.,
Gustavo Freire
"""

# formatação dos números (moeda, dinheiro)

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# enviar o e-mail
pyautogui.hotkey("ctrl", "enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[51]:


time.sleep(5)
pyautogui.position()


# In[ ]:




