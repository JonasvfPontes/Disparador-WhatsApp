#bibliotecas necessárias
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import urllib

#Criar validador de mensalidade

#Lendo contatos
df_contatos = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Lista Contatos', usecols="A:B", engine='openpyxl')
#print(df_contatos)

#Lendo Mensagem e transformando em String
df_mensagem = pd.read_excel('Texto e Lista de contatos.xlsx', sheet_name='Texto', usecols="A", engine='openpyxl')
mensagem = str(df_mensagem.loc[0, 'Escreva o texto no campo abaixo']) #Pegando primeira linha da coluna 'Escrev o texto...'

#Criando Objeto Navegador
navegador = webdriver.Chrome()
navegador.get('https://web.whatsapp.com/')

#Abrindo e aguardando login no WhatsApp
while len(navegador.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)

#Mandar mensaggem para cada contato
for i, nome in enumerate(df_contatos['Primeiro Nome']):
    if len(str(df_contatos.loc[i, 'Numero'])) >= 11:
        numero = str(int(df_contatos.loc[i,'Numero']))
        mensagemPersonalizada = urllib.parse.quote(mensagem.replace('--CONTATO--', nome))
        link = f'https://web.whatsapp.com/send?phone={numero}&text={mensagemPersonalizada}'
        navegador.get(link)
        while len(navegador.find_elements(By.ID, 'side')) < 1:
            time.sleep(5)
        #Procurar  xPATH, se error esperar até 5 segundos, senão ir para o próximo contato
        cont=0 #Contador de espera
        while cont <5:
            try:
                navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                time.sleep(1)
                navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div/p/span').send_keys(Keys.ENTER)
                break
            except:
                #Esperar mais um pouco
                time.sleep(1)
                #print(cont)
                cont = cont + 1
        time.sleep(2) #Tempo depois que enviar cada mensagem
navegador.close