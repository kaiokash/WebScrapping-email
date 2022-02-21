from sys import prefix
import pyautogui  
import time 
import pyperclip
import _tkinter
import pandas
import openpyxl

from selenium import webdriver
from selenium.webdriver.common.keys import Keys

import smtplib  

import email.message 

#deixar invisivel
#options = webdriver.ChromeOptions()
#options.add_argument("--headless")

#nav = webdriver.Chrome(chrome_options=options)

nav = webdriver.Chrome("chromedriver.exe")

#Pesquisar o site a cotação '''
nav.get("https://www.google.com/")

nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação dólar")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_dolar = nav.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")

print(cotacao_dolar)  



    
# Passo 2: Pegar a cotação do Euro
nav.get("https://www.google.com/")

nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("cotação euro")
nav.find_element_by_xpath('/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)

cotacao_euro = nav.find_element_by_xpath('//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value=1')
    
    
    
    
   # print(cotacao_euro)

print(cotacao_euro)

#abrir em nova aba sem click

#aba1 = nav.window_handles[0]

       
#nav.execute_script("window.open()")

# Switch to the newly opened tab
#nav.switch_to.window(nav.window_handles[1])

# Navigate to new URL in new tab

# Run other commands in the new tab here



# Passo 3: Pegar a cotação do Ouro
nav.get("https://www.melhorcambio.com/ouro-hoje")


cotacao_ouro = nav.find_element_by_id('comercial').get_attribute('value')



cotacao_ouro = cotacao_ouro.replace("," , ".")
#round(cotacao_ouro)

print(cotacao_ouro) 






nav.quit()
# Passo 4: Importar a lista de produtos
import pandas as pd

tabela = pd.read_excel("Produtos.xlsx")
print (tabela)


# Passo 5: Recalcular o preço de cada produto
# atualizar a cotação
# nas linhas onde na coluna "Moeda" = Dólar
tabela.loc[tabela["Moeda"] == "Dólar", "Cotação"] = float(cotacao_dolar)
tabela.loc[tabela["Moeda"] == "Euro", "Cotação"] = float(cotacao_euro)
tabela.loc[tabela["Moeda"] == "Ouro", "Cotação"] = float(cotacao_ouro)




# atualizar o preço base reais (preço base original * cotação)
tabela["Preço Base Reais"] = tabela["Preço Base Original"] * tabela["Cotação"]

# atualizar o preço final (preço base reais * Margem)
tabela["Preço Final"] = tabela["Preço Base Reais"] * tabela["Margem"]

# tabela["Preço de Venda"] = tabela["Preço de Venda"].map("R${:.2f}".format)

print(tabela)


# Passo 6: Salvar os novos preços dos produtos
tabela.to_excel("Produtos Novo.xlsx", index=False)



print(tabela)

dol = float(cotacao_dolar)
dol=round(dol,2)
print(dol)

eur = float(cotacao_euro)
eur=round(eur,2)
print(eur)

our = float(cotacao_ouro)
our=round(our,2)
print(our)



def enviar_email():
    corpo_email = f"""
    <p> Bom dia!</p>
    <p>Testando aqui... </p>
    <p>Cotações de hj..:\n</p>

    <p>Dólar: {dol} R$</p>
    <p>Euro:  {eur} R$</p>
    <p>Ouro:  {our} R$</p>

    """
    msg = email.message.Message()
    msg['Subject'] = "Email automatico Teste"
    msg['From'] = 'teste@teste.com'
    msg['To'] = 'teste2@teste.com'
    password = 'teste@123'
    msg.add_header('Content-Type','text/html')
    msg.set_payload(corpo_email)



    s = smtplib.SMTP('smtp.gmail.com:587')
     

    s.ehlo
    s.starttls()
    #Credenciais
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.   as_string().encode('utf-8'))



    print('Email Enviado')
   
enviar_email()
   

