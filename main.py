#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import *
from pgAvulso import *
from downloads import direcionarDownloads, opçoesdoChrome
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime

#data
today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")
print(data_em_texto)
#preferencias do chrome
chrome_options = webdriver.ChromeOptions()
opçoesdoChrome
driver = Chrome(chrome_options=chrome_options)

#criar a pasta do dia no onedrive
criarPasta(driver, data_em_texto)

#Inicia o navegador
chamarDriver(driver)

#Login

fazerLogin(driver)

#clicando no modulo financeiro
#direcionaSolicitado(driver)
#tramitando
#tramitar(driver)

#Modulo Pagamento Avulso

pgtoAvulso(driver)
chamarSharepoint(driver, Keys)
baixarNf(driver)
