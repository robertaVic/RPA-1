#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes2 import chamarDriver, fazerLogin, criarPasta, direcionaSolicitado
from pgAvulso2 import pgtoAvulso, chamarSharepoint, baixarNf
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime

#data
today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")
print(data_em_texto)
#preferencias do chrome
chrome_options = webdriver.ChromeOptions()
#permitir notificaçoes e download automatico
prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1,
        "download.default_directory": r"C:\Users\Usuario\OneDrive - tpfe.com.br\{data_em_texto}\{pastaSolicitacao}"}
        
chrome_options.add_experimental_option("prefs",prefs)
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
