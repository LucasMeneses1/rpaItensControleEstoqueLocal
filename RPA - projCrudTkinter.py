#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyautogui
import pandas as pd
import time

df = pd.read_excel('itens.xlsx')

def preencher_campo(img, txt):
    x , y, largura, altura = pyautogui.locateOnScreen(f'{img}.png')
    pyautogui.click(x + largura/2, y + altura/2)
    pyautogui.write(txt)

def clicar_btn(img):
    x , y, largura, altura = pyautogui.locateOnScreen(f'{img}.png')
    pyautogui.click(x + largura/2, y + altura/2)
        

pyautogui.alert('O código vai começar. Não mexa em NADA enquanto o código tiver rodando. Quando finalizar, eu te aviso')

x , y, altura, largura = pyautogui.locateOnScreen('projetocrud.png')
pyautogui.click(x + largura/2, y + altura/2, 2, 0.1)

while not pyautogui.locateOnScreen('login.png'):
    time.sleep(1)
    
preencher_campo('login', 'usuario1')
preencher_campo('senha', 'senha1')

clicar_btn('entrar')

while not pyautogui.locateOnScreen('criar.png'):
    time.sleep(1)
    
clicar_btn('criar')

j = len(df['cod'])
for i in range(j):
    cod = df.iloc[i, 0]
    nome = df.iloc[i, 1]
    lote = df.iloc[i, 2]
    qtd = df.iloc[i, 3]
    print(cod, nome, lote, qtd)
    preencher_campo('cod', cod)
    time.sleep(1)
    preencher_campo('nome', nome)
    time.sleep(1)
    preencher_campo('lote', str(lote))
    time.sleep(1)
    preencher_campo('qtd', str(qtd))
    time.sleep(1)

    clicar_btn('confirmar')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)

