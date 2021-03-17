#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import chamarDriver, fazerLogin, direcionaSolicitado
from pgAvulso import pgtoAvulso
from selenium.webdriver.chrome.options import Options
#preferencias do chrome
chrome_options = webdriver.ChromeOptions()
#permitir notificaçoes e download automatico
prefs = {"profile.default_content_setting_values.notifications" : 1, 'profile.default_content_setting_values.automatic_downloads': 1}
chrome_options.add_experimental_option("prefs",prefs)
driver = Chrome(chrome_options=chrome_options)




<<<<<<< HEAD
#paulo é ruim e chato
=======
#paulo é ruim
#TESTE3
>>>>>>> c54bc918f3d2c0496283494b57dad64e1f37de93

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
