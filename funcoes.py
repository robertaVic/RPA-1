from time import sleep, time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from gerenciadorPastas import recuperar_diretorio_usuario
from selenium.webdriver.chrome.options import Options

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def espera_explicita_de_elemento(drive, element_path, acao, informacao_acao, tempo_espera):
    
    if acao == "encontrar":
        try:
            element = WebDriverWait(drive, tempo_espera).until(
                EC.presence_of_element_located((By.XPATH, element_path))
            )
            sleep(1)
        
            return False
        except:
            print(informacao_acao, "Erro ao encontrar elemento")
            return True


    elif acao == "click":
        try:
            element = WebDriverWait(drive, tempo_espera).until(
                EC.element_to_be_clickable((By.XPATH, element_path))
            )
            #print("Encontrou o bendito")
            sleep(1)
            element.click()
            return False
            #drive.find_element_by_xpath(element_path).click()
        except:
            print(informacao_acao, "Erro ao encontrar elemento")
            return True

#preferencias do chrome
def padraoChrome(diretorio):
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1, "download.default_directory": recuperar_diretorio_usuario() + diretorio }
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options

#Inicia o navegador
def chamarDriver(navegador):
    navegador.maximize_window()
    #Limpando a cache
    builder = ActionChains(navegador)
    builder.key_down(Keys.CONTROL).key_down(Keys.F5)
    builder.key_up(Keys.CONTROL).key_up(Keys.F5)
    builder.perform()

    navegador.get("https://tpf2.madrix.app/")
    
    #navegador.implicitly_wait(10)

# #Faz login   
def fazerLogin(login):
    espera_explicita_de_elemento(login, "/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[1]/div/div/div/input", "click", "Login", 120)
    login.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[1]/div/div/div/input").send_keys("roberta.costa")
    espera_explicita_de_elemento(login, "/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[1]/div/div/div/input", "click", "Senha", 120)
    login.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[2]/div/div/div/input").send_keys("123")
    # espera_explicita_de_elemento(login, "/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[3]/div/button/span[1]", "click", "Login", 30)
    espera_explicita_de_elemento(login, "/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[3]/div/button/span[1]", "click", "Login", 120)

def encontrar_elemento_por_repeticao(drive, element_path, acao, informacao_acao, tempo_espera):
    maximo_tentativas = 0
    while maximo_tentativas <= 20:
        print(informacao_acao, maximo_tentativas)
        try:
            sleep(tempo_espera)
            drive.find_element_by_xpath(element_path)
            if acao == "click":
                drive.find_element_by_xpath(element_path).click()
                maximo_tentativas = 21
            elif acao == "link":
                maximo_tentativas = 21
                pass
        except:
            maximo_tentativas+=1
            sleep(tempo_espera)
    if maximo_tentativas > 20:
        return("#Erro " + informacao_acao)

'''
page = financeiro.current_window_handle
for handle in financeiro.window_handles: 
        if handle != page: 
            lpage = handle 
    financeiro.switch_to.window(lpage)
'''
#financeiro.execute_script("document.getElementByXpath('/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]')[0].click()")

