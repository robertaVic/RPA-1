#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import *
from pgAvulso import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime

def direcionarDownloads(id):
    #
    global pref
    pref = (r"C:\\Users\\Usuario\\OneDrive - tpfe.com.br\\" + str(data_em_texto) + "\\"+ str(id))
#data
today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")
print(data_em_texto)
print(pref)
#preferencias do chrome
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1, "download.default_directory": pref }
chrome_options.add_experimental_option("prefs", prefs)
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
