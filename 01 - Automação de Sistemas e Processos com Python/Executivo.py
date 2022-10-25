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

# In[73]:


pip install pyautogui
pip install pyperclip


# In[74]:


import pyautogui
import pyperclip
import time

# 1° logar no sistema da empresa (link do drive)
# pyautogui.hotkey para atalhos | .write = escrever um texto | .press apertar 1 tecla

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl","t")
#pyautogui.press('win')
#pyautogui.write("chrome")
#pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")

# agurda um tempo setado.
time.sleep(5)

# 2° navegar no sistema e eontrar os dados (entrar na pasta exportar) clicks ="right"
pyautogui.click(x=353, y=269, clicks=2)
time.sleep(2)

# 3° exportar/fazer o donwload do DB.
pyautogui.click(x=432, y=262)
pyautogui.click(x=1221, y=156)
pyautogui.click(x=969, y=600)
#aguardar o fim do download
time.sleep(60)


# ### Vamos agora ler o arquivo baixado para pegar os indicadores
# 
# - Faturamento
# - Quantidade de Produtos

# In[75]:


# 4° importar o DB para p PY
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Administrator\Downloads\Vendas - Dez.xlsx")
#ou usar o print em outro compilador | para selecionar ou mais abas colocar sheets =2
display(tabela)


# In[76]:


# 5° Calcular os indicadores

#faturamento é a soma das vendas
faturamento = tabela["Valor Final"].sum()
print(faturamento)

#quantidade de produtos
quantidade = tabela["Quantidade"].sum()
print(quantidade)


# ### Vamos agora enviar um e-mail pelo gmail

# In[77]:


#fechar icone de downloads
pyautogui.click(x=1345, y=696)


# In[78]:



# 6° enviar o email com o relatorio para a diretoria
pyautogui.hotkey("ctrl","t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey("ctrl","v")
pyautogui.press("enter")


# agurda um tempo setado.
time.sleep(5)

#aperte o escrever
pyautogui.click(x=150, y=156)
time.sleep(2)
#escrever destinatrio e email
pyautogui.click(x=871, y=316)
pyautogui.write("tamer.srhn@hotmail.com")
pyautogui.press("tab")
time.sleep(1)
# escrever tema | pode selecionar ou apertar o tab
pyautogui.click(x=869, y=348, clicks = 2)
pyperclip.copy("Relatórios de Vendas")
pyautogui.hotkey("ctrl", "v")
time.sleep(1)
#escrevendo o conteudo
pyautogui.click(x=829, y=391)
texto = f"""
Prezados Bom Dia.

O faturamento da empresa na data de ontem foi de R${faturamento:,.2f}
a quantidade de produtos foi de {quantidade:,}

att.
Tamer Serhan"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.click(x=962, y=698)

#selecionar arquivo
pyautogui.click(x=355, y=50)
pyautogui.write(r"C:\Users\Administrator\Downloads")
pyautogui.press("enter")
pyautogui.click(x=252, y=177, clicks =2)

#envio do email
pyautogui.hotkey("ctrl","enter")


# #### Use esse código para descobrir qual a posição de um item que queira clicar
# 
# - Lembre-se: a posição na sua tela é diferente da posição na minha tela

# In[79]:


time.sleep(5)
pyautogui.position()


# In[ ]:




