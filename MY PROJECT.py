#!/usr/bin/env python
# coding: utf-8

# ### Análise de Dados com Python
# 
# Você trabalha em uma empresa de telecom e tem clientes de vários serviços diferentes, entre os principais: internet e telefone.
# 
# O problema é que, analisando o histórico dos clientes dos últimos anos, você percebeu que a empresa está com Churn de mais de 26% dos clientes.
# 
# Isso representa uma perda de milhões para a empresa.
# 
# O que a empresa precisa fazer para resolver isso?
# 
# - acessar base de dados (.csv) na internet (selenium) e baixar
#     - link: https://drive.google.com/drive/folders/1T7D0BlWkNuy_MDpUHuBG44kT80EmRYIs?usp=sharing
# - enviar o arquivo para a pasta do projeto (pathlib, shutil)
# - acessar o df (pandas)
# - análise desejada (pandas, plotly, matplotlib...)
# - exportar os resultados para um arquivo .docx existente (python-docx)
# - criar uma pasta nova, na pasta do projeto, e adicionar os gráficos (pathlib, shutil)
# - enviar o arquivo, com os gráficos em anexo, por email (win32com.client)
# 

# In[1]:


#1 - importar bibliotecas
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
servico = Service(ChromeDriverManager().install())

import pyautogui
pyautogui.PAUSE = 1
import time
import os
import shutil
import pandas as pd
import plotly.express as px
from docx import Document
import win32com.client as win32
from datetime import datetime


# In[ ]:


#2 - acessar a base de dados na internet
browser = webdriver.Chrome()
browser.maximize_window()
browser.get('https://drive.google.com/drive/folders/1T7D0BlWkNuy_MDpUHuBG44kT80EmRYIs?usp=sharing')

while not len(browser.find_elements(By.XPATH, '//*[@id="gb"]/div[2]/div[1]/div[4]/div/a/img')) > 0 :
    time.sleep(1)
print('ok')

while not pyautogui.locateOnScreen('Screenshot_2.png', grayscale=True, confidence=0.9):
    time.sleep(1)
print('ok')

imagem = pyautogui.locateOnScreen('Screenshot_2.png', grayscale=True, confidence=0.9)
pyautogui.click(pyautogui.center(imagem))

pyautogui.click(x=1833, y=147)
time.sleep(5)
browser.quit()


# In[8]:


#3 - enviar o arquivo para a pasta do projeto
local_arquivo = 'C:/Users/W10/Downloads/telecom_users.csv'
local_destino = os.getcwd()

shutil.copy2(local_arquivo, local_destino)
os.remove(local_arquivo)


# In[9]:


#4 - acessar a base de dados
df = pd.read_csv('telecom_users.csv', sep=',')
pd.set_option('display.max_columns', None)
display(df)
df.info()


# In[10]:


#5 - tratamento do df e dos dados
df = df.drop('Unnamed: 0', axis=1)
df = df.dropna(axis=1, how='all')
df = df.dropna(axis=0, how='any')
df['TotalGasto'] = pd.to_numeric(df['TotalGasto'], errors='coerce')
display(df)
df.info()


# In[11]:


#6.1 - análise: quais os números dos cancelamentos/churns?
cancelamentos = df['Churn'].value_counts()
df_churn = pd.DataFrame(cancelamentos)
display(df_churn)

cancelamentos_perc = df['Churn'].value_counts(normalize=True).map('{:.2%}'.format)
df_churn_perc = pd.DataFrame(cancelamentos_perc)
display(df_churn_perc)

clientes = len(df)
clientes_cancelamentos = df_churn_perc.loc['Sim', 'Churn']
clientes_cancelamentos2 = df_churn.loc['Sim', 'Churn']
print(f'Número de clientes: {clientes:,}; Número de cancelamentos: {clientes_cancelamentos2:,}')


# In[12]:


#6.2 - análise: quais os clientes mais fiéis (maior tempo como cliente)? quantos destes clientes cancelaram?
max_clientes = df['MesesComoCliente'].sort_values(ascending=False)
df_max_clientes = pd.DataFrame(max_clientes)
df_max_clientes = df_max_clientes.reset_index()
df_max_clientes = df_max_clientes.rename(columns={'index':'ID Cliente'})
df_max_clientes = df_max_clientes.drop(0, axis=0)
display(df_max_clientes[:307])

print(f'Lista de clientes (IDs) mais fiéis:')
df_max_clientes_top = df_max_clientes['ID Cliente'][:307]
for cliente in df_max_clientes_top:
    print(f'{cliente}', end=', ')

    
fig_max_clientes_top = px.histogram(df, x='MesesComoCliente', color='Churn', title=f'Comparação entre os clientes mais fiéis e Churn', text_auto=True)
fig_max_clientes_top.update_traces(textfont_size=20)
fig_max_clientes_top.show()


# In[13]:


#6.3 - análise: comparação da quantidade de clientes para cada forma de pagamento; quantos de cada cancelou?
pagamento = df['FormaPagamento'].value_counts()
df_pagamento = pd.DataFrame(pagamento)
display(df_pagamento)

fig = px.pie(df_pagamento, values=df_pagamento['FormaPagamento'], names=df_pagamento.index)
fig.update_traces(textfont_size=25)
fig.show()

fig_pgto = px.histogram(df, x='FormaPagamento', color='Churn', title='Comparação entre as formas de pagamento mais fiéis e Churn', text_auto=True)
fig_pgto.update_traces(textfont_size=20)
fig_pgto.show()


# In[14]:


#6.5 - análise: qual o faturamento mensal da empresa?
df_fat = df[['ValorMensal', 'Churn']]
df_fat = df_fat.groupby('Churn').sum()
display(df_fat)

faturamento_mensal = df_fat.loc['Nao', 'ValorMensal']
print(f'Faturamento mensal = R$ {faturamento_mensal:,.2f}.')


# In[15]:


#6.6 - análise: quantos clientes possuem o pacote completo da empresa (ServicoTelefone	MultiplasLinhas	ServicoInternet	ServicoSegurancaOnline	ServicoBackupOnline	ProtecaoEquipamento	ServicoSuporteTecnico	ServicoStreamingTV	ServicoFilmes)? Destes, quais cancelaram?
contador = 0
for linha in df.index:
    if df.loc[linha, 'ServicoTelefone'] == 'Sim' and df.loc[linha, 'MultiplasLinhas'] == 'Sim' and df.loc[linha, 'ServicoInternet'] != 'Não' and df.loc[linha, 'ServicoSegurancaOnline'] == 'Sim' and df.loc[linha, 'ServicoBackupOnline'] == 'Sim' and df.loc[linha, 'ProtecaoEquipamento'] == 'Sim' and df.loc[linha, 'ServicoSuporteTecnico'] == 'Sim' and df.loc[linha, 'ServicoStreamingTV'] == 'Sim' and df.loc[linha, 'ServicoFilmes'] == 'Sim':
        contador += 1
        
print(f'\n\nQuantidade de clientes com o pacote completo: {contador}.')

lista_itens_pacote_completo = ['ServicoTelefone', 'MultiplasLinhas', 'ServicoInternet', 'ServicoSegurancaOnline', 'ServicoBackupOnline', 'ProtecaoEquipamento', 'ServicoSuporteTecnico', 'ServicoStreamingTV', 'ServicoFilmes']
for item in lista_itens_pacote_completo:
    fig_pacotecompleto = px.histogram(df, x=item, color='Churn', title=f'Comparação entre clientes que tenham {item} e Churn', text_auto=True)
    fig_pacotecompleto.update_traces(textfont_size=20)
    fig_pacotecompleto.show()


# In[16]:


#6.7 - análise: qual o valor médio dos contratos mensais? e dos contratos anuais? quantos destes clientes cancelaram?
df_media = df.groupby('TipoContrato').mean()
df_media = df_media['ValorMensal']
df_media = df_media.reset_index()
df_media = df_media.rename(columns={'ValorMensal':'MédiaValorMensal'})
df_media['MédiaValorMensal'] = df_media['MédiaValorMensal']
display(df_media)

mediatotal = df_media['MédiaValorMensal'].mean()
print(f'Valor médio de todos os contratos R$ {mediatotal:,.2f}.')

fig_contrato = px.histogram(df, x='TipoContrato', color='Churn', title='Comparação entre os tipos de contrato e Churn', text_auto=True)
fig_contrato.update_traces(textfont_size=20)
fig_contrato.show()


# In[17]:


#6.8 - análise: qual o percentual de clientes que tem o plano básico (internet fibra e telefone)?
contador2 = 0
for linha in df.index:
    if df.loc[linha, 'ServicoTelefone'] == 'Sim' and df.loc[linha, 'ServicoInternet'] != 'Não':
        contador2 += 1

percentual_internetfibra_telefone = contador2 / len(df)
print(f'O percentual de clientes que tem o plano básico é {percentual_internetfibra_telefone:.0%}.')


# In[20]:


#6.9 - análise: quais os reais motivos dos cancelamentos/churns?
for col in df.columns:
    fig = px.histogram(df, x=col, color=df['Churn'], title=f'Comparação entre {col} e Churn', text_auto=True)
    fig.update_traces(textfont_size=20)
    fig.show()
    # download do gráfico
    fig.write_image(f"{col}-churn.png")


# In[19]:


#7 - exportar os resultados das análises para um arquivo de texto existente .docx
documento = Document('Relatório Mensal Análise de Dados.docx')

relatorio = '''De acordo com a análise de dados, os principais motivos encontrados para os churns do mês são:

Clientes solteiros,
Clientes sem dependentes,
Clientes com poucos meses de contrato,
Clientes com serviço telefônico,
Clientes com serviço de internet fibra,
Clientes sem serviço de segurança online,
Clientes sem serviço de backup online,
Clientes sem serviço de proteção de equipamento,
Clientes sem serviço de suporte técnico,
Clientes com contrato mensal,
Clientes com fatura digital,
Clientes com forma de pagamento por meio de boleto eletrônico.
'''    
    
dicionario = {
    'item1': str(f'{clientes:,}'),
    'item2': str(f'{clientes_cancelamentos2:,}'),
    'item3': str(clientes_cancelamentos),
    'item4': str(len(df_max_clientes_top)),
    'item5': str(f'R$ {faturamento_mensal:,.2f}'),
    'item6': str(contador),
    'item7': str(f'{contador2:,}'),
    'item8': str(f'R$ {mediatotal:,.2f}'),
    'item9': str(relatorio)
}

for paragrafo in documento.paragraphs:
    for chave, valor in dicionario.items():
        paragrafo.text = paragrafo.text.replace(chave, valor)

documento.save(fr'{local_destino}/Relatório Mensal Análise de Dados (FINAL).docx')


# In[25]:


#8 - criar uma pasta nova, na pasta do projeto, e adicionar os gráficos .png
try:
    os.mkdir(fr'{local_destino}/Gráficos')
except:
    pass

for arquivo in os.listdir(local_destino):
    if 'churn.png' in arquivo:
        grafico = fr'{local_destino}/{arquivo}'
        pasta_graficos = fr'{local_destino}/Gráficos'
        shutil.move(grafico, pasta_graficos)


# In[26]:


#9 - enviar o arquivo de texto, com os gráficos em anexo, por email para a diretoria
hoje = datetime.now().strftime('%d/%m/%Y')

graficos = fr'{local_destino}/Gráficos'
lista_graficos = []
for arquivo in os.listdir(graficos):
    lista_graficos.append(arquivo)

outlook = win32.Dispatch('Outlook.Application')

email = outlook.CreateItem(0)
email.To = 'bep_rafael@hotmail.com'
email.Subject = 'Relatório mensal de análise de dados - Churn'
email.Body = f'''
Olá Diretoria TELECOM,

Seguem anexados um arquivo de texto e gráficos da análise de dados a partir de nossa base de dados de clientes TELECOM.
Data de envio deste e-mail: {hoje}.

Atenciosamente,
Analista de dados TELECOM.
'''   
for grafico in lista_graficos:
    email.Attachments.Add(fr'{local_destino}/Gráficos/{grafico}')
email.Attachments.Add(fr'{local_destino}/Relatório Mensal Análise de Dados (FINAL).docx')
email.Save()
email.Send()


# In[ ]:




