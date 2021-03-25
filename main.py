#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime
from gerenciadorPastas import criarPastaData
from pgAvulso import pagamentoAvulso
import getpass


#preferencias do chrome
def padraoChrome(diretorio):
    usuario = getpass.getuser()
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1, "default_directory": "C:\\Users\\"+ usuario +"\\" + diretorio }
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options

diretorioPadrao = "Downloads"
chrome_options = padraoChrome(diretorioPadrao)

usuario = getpass.getuser()
print(usuario)

today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")
pref = "C:\\Users\\"+ usuario +"\\OneDrive - tpfe.com.br\\RPA-DEV" 

driver = Chrome(chrome_options=chrome_options)

criarPastaData(pref, data_em_texto)
#criar a pasta do dia no onedrive
#criarPasta(driver, data_em_texto)

#Inicia o navegador
chamarDriver(driver)

#Login
fazerLogin(driver)

#entra no pagamento avulso
pagamentoAvulso(driver)

#volta para o sgp para baixar as NF'S
# baixarNf(driver)
