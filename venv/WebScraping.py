import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import pygetwindow
import pywinauto.mouse as mouse
from pywinauto.keyboard import send_keys
import pandas as pd
from pandas import DataFrame
import numpy as np
import matplotlib.pyplot as plt
import sys
import re
import six
import requests
from bs4 import BeautifulSoup
#import html5lib
import lxml
from collections import Counter
from openpyxl import Workbook
import xlrd


foro = pd.read_excel(r'C:\Users\mac1218\PycharmProjects\pythonProject5\venv\Pasta2.xlsx', sheet_name='Planilha1', header=0, usecols="B")

nome_do_arquivo = 'Pasta2.xlsx'
path = nome_do_arquivo
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
m_row = sheet_obj.max_row
for i in range(1, m_row + 1):
    def remove(i, m_row):
        if i != None:
            return
i.delete_rows(m_row[0].row, 1)

cell_obj = sheet_obj.cell(row=i, column=2)

# processo = cell_obj.value.split('.8.26.')
# cell_obj.value = ','.join(processo)
print(cell_obj.value)

for index,row in foro.iterrows():
 print(row["Processo"])
 break

#if (row["Processo"] == 0):


#print(str(index) + row["Processo"])

#print(row['Processo'].split('.8.26.'))
Processo = row["Processo"].split('.8.26.')

   #print(Processo[0])
print(Processo)

#options = webdriver.ChromeOptions()
#options.headless = False
#options.add_extension(r'C:\Users\mac1017\Desktop\BaseDadosInicial\ChromeDrive\extensão\Web Signer 2.14.3.0.crx')
navegador = webdriver.Chrome(r"C:/Users/mac1218/DBRobo_Inicial/SP/DriveChrome/chromedriver.exe")
# navegador = webdriver.Chrome(r"C:/Users/mac1017/Desktop/Robôs_finalizados/DriveChrome/chromedriver.exe")
navegador.implicitly_wait(20)

navegador.get("https://esaj.tjsp.jus.br/esaj/portal.do?servico=190090")
navegador.find_element_by_xpath('//*[@id="identificacao"]/strong/a').click()
# time.sleep(5)
# Clica no cpf
navegador.find_element_by_xpath('//*[@id="linkAbaCpf"]').click()
time.sleep(5)
# INSERE O LOGIN
navegador.find_element_by_xpath('/html/body/table[4]/tbody/tr/td/table[2]/tbody/tr[1]/td[1]/div/table/tbody/tr[1]/td/div[2]/div[2]/form/div[1]/table/tbody/tr[1]/td[2]/input').send_keys('10208311890')
time.sleep(5)
# insere a senha
navegador.find_element_by_xpath('/html/body/table[4]/tbody/tr/td/table[2]/tbody/tr[1]/td[1]/div/table/tbody/tr[1]/td/div[2]/div[2]/form/div[1]/table/tbody/tr[2]/td[2]/input').send_keys('mac@abril')
time.sleep(5)

# clica no botão
navegador.find_element_by_xpath('/html/body/table[4]/tbody/tr/td/table[2]/tbody/tr[1]/td[1]/div/table/tbody/tr[1]/td/div[2]/div[2]/form/div[1]/table/tbody/tr[4]/td[2]/input[4]').click()
# time.sleep(50)
#cLICA EM CONSULTA PRIMEIRO GRAU
navegador.find_element_by_xpath('/html/body/table[3]/tbody/tr/td[1]/ul/li[2]/ul/li[1]/a').click()
time.sleep(3)
#Clica em inserir processo
navegador.find_element_by_xpath('//*[@id="numeroDigitoAnoUnificado"]').send_keys(foro)
time.sleep(2)
#Clica no segundo campo para inserir o processo
#navegador.find_element_by_xpath('//*[@id="foroNumeroUnificado"]').send_keys(Processo[1])
#clica no botão consultar
time.sleep(5)
navegador.find_element_by_xpath('//*[@id="botaoConsultarProcessos"]').click()
time.sleep(5)

#Move a barra de rolagem
navegador.execute_script('window.scrollTo(0, window.scrollY + 200)')
time.sleep(7)


navegador.find_element_by_xpath('/html/body/div[2]/table[2]').click()
element = navegador.find_element_by_xpath('//*[@id="tabelaUltimasMovimentacoes"]')
html_content = element.get_attribute('outerHTML')

#print(html_content)

soup = BeautifulSoup(html_content, 'lxml')

table = soup.find_all('td', class_='descricaoMovimentacao')[1]

print(table.text)

#Grava no arquivo txt a publicação a ser extraida
arquivo = open('arq01.txt','w')
arquivo.write(table.text)
arquivo.close()

#faz a busca por palavra chave e retorna a linha da palavra encontrada
regex_palavras = re.compile(r'\b(\w+)\b')
# palavra a ser buscada
busca = input('Digite a palavra desejada: ')

with open('arq01.txt', 'r') as arquivo:
    for linha in arquivo:  # assume que cada linha do arquivo é uma frase
        # buscar as palavras da linha, usando a regex
        for match in regex_palavras.finditer(linha.strip()):
            if busca == match.group(1): # group(1) contém a palavra
                print(linha)
                break




#Grava na planilha do excel
#grava = pd.DataFrame({'Descrição': {linha}})
# Determina o caminho da planilha
#file_name = (r'C:\Users\mac1218\TESTE\Pasta2.xlsx')

#Salvando na planinha excel
#foro.to_excel(file_name)
#print('DataFrame is written to Excel File successfully.')





#df_full = pd.read_html(table.text)
#df = df_full['fundoClaro containerMovimentacao: 0','dataMovimentacao','descricaoMovimentacao']

#print(df_full)
#Fecha somente a janela do navegador atual
handles = navegador.window_handles

for i in handles:
    navegador.switch_to.window(i)

    if navegador.title == "Portal de Serviços e-SAJ":
        time.sleep(2)
        navegador.quit()
#sys.exit()












#Pega a variavel navegador do selenium já logado e faz a webscraping da página
#html = navegador.page_source
#soup = BeautifulSoup(html, 'lxml')
#lista = soup.find('tr', {'class':'fundoClaro containerMovimentacao'}) pega ultimo andamento
#Pega o texto da página
#lista1 = soup.find('span', {'style':'font-style: italic;'})
#lista1 = soup.find_all('tr', {'style':'margin-left:15px; margin-top:1px;'})

#print(lista1)

#for i in lista1:
#   print(i.text)

#for i in range(10):
# filhas = i.find_all("td")
# print(filhas[0])
# print(filhas[1])
# print(filhas[2])


#print(lista1)

#Grava na planilha do excel
#foro = pd.DataFrame({'Descrição': {lista1}})
# Determina o caminho da planilha
#file_name = (r'C:\Users\mac1218\TESTE\Pasta2.xlsx')

#Salvando na planinha excel
#foro.to_excel(file_name)
#print('DataFrame is written to Excel File successfully.')

#teste


