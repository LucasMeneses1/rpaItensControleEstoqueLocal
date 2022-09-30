#!/usr/bin/env python
# coding: utf-8

# In[3]:


from selenium import webdriver
import time
import pandas as pd
from selenium.webdriver.common.by import By

driver = webdriver.Chrome()

link = "http://127.0.0.1:5000"
driver.get(link)

driver.find_element(By.ID, 'usuario').send_keys('usuario1')
driver.find_element(By.ID, 'senha').send_keys('senha1')
driver.find_element(By.ID, 'logar').click()
time.sleep(1)
lista_itens = pd.read_excel("Itens.xlsx")
driver.find_element(By.XPATH, '/html/body/div[3]/form/a[1]/button').click()
time.sleep(1)
j = len(lista_itens['cod'])
for i in range(j):
    cod = lista_itens.iloc[i, 0]
    nome = lista_itens.iloc[i, 1]
    lote = lista_itens.iloc[i, 2]
    qtd = lista_itens.iloc[i, 3]
    
    driver.find_element(By.ID, 'cod').send_keys(cod)
    driver.find_element(By.ID, 'nome').send_keys(nome)
    driver.find_element(By.ID, 'lote').send_keys(str(lote))
    driver.find_element(By.ID, 'qtd').send_keys(int(qtd))
    driver.find_element(By.ID, 'confirmar').click()
    time.sleep(1)
    driver.find_element(By.ID, 'qtd').clear()


# In[2]:


driver.find_element(By.ID, 'usuario').send_keys('usuario1')
driver.find_element(By.ID, 'senha').send_keys('senha1')
driver.find_element(By.ID, 'logar').click()


# In[5]:


time.sleep(1)
lista_itens = pd.read_excel("Itens.xlsx")
lista_itens.iloc[0, 0]


# In[4]:


driver.find_element(By.ID, 'qtd').clear()


# In[8]:


driver.find_element(By.XPATH, '/html/body/div[5]/div/a[1]/div/button').click()
time.sleep(1)
j = len(lista_itens['cod'])
for i in range(j):
    cod = lista_itens.iloc[i, 0]
    nome = lista_itens.iloc[i, 1]
    lote = lista_itens.iloc[i, 2]
    qtd = lista_itens.iloc[i, 3]
    
    driver.find_element(By.ID, 'cod').send_keys(cod)
    driver.find_element(By.ID, 'nome').send_keys(nome)
    driver.find_element(By.ID, 'lote').send_keys(str(lote))
    driver.find_element(By.ID, 'qtd').send_keys(int(qtd))
    driver.find_element(By.ID, 'confirmar').click()
    time.sleep(1)
    driver.find_element(By.ID, 'qtd').clear()
    
    

