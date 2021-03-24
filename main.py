#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import *
from pgAvulso import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime


#preferencias do chrome
def padraoChrome(diretorio):
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1, "default_directory": "C:\\Users\\Usuario\\" + diretorio }
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options


diretorioPadrao = "Downloads" 
chrome_options = padraoChrome(diretorioPadrao)
driver = Chrome(chrome_options=chrome_options)

#criar a pasta do dia no onedrive
#criarPasta(driver, data_em_texto)

#Inicia o navegador
chamarDriver(driver)

#Login
fazerLogin(driver)

#entra no pagamento avulso
pgtoAvulso(driver)

#volta para o sgp para baixar as NF'S
# baixarNf(driver)
