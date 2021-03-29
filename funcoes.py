from time import sleep
from datetime import date
from selenium import webdriver
from gerenciadorPastas import recuperar_diretorio_usuario
from selenium.webdriver.chrome.options import Options

#preferencias do chrome
def padraoChrome(diretorio):
    chrome_options = webdriver.ChromeOptions()
    prefs = {"profile.default_content_setting_values.notifications" : 1, "profile.default_content_setting_values.automatic_downloads": 1, "download.default_directory": recuperar_diretorio_usuario() + diretorio }
    chrome_options.add_experimental_option("prefs", prefs)
    return chrome_options

#Inicia o navegador
def chamarDriver(navegador):
    navegador.get("https://tpf.madrix.app/?next=/app")
    navegador.maximize_window()
    navegador.implicitly_wait(100)

# #Faz login   
def fazerLogin(login):
    login.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[1]/div/div/div/input").send_keys("roberta.costa")
    login.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[2]/div/div/div/input").send_keys("123")
    login.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[3]/div/button/span[1]").click()

def direcionaSolicitado(financeiro):
    financeiro.implicitly_wait(120)
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div[1]").click()
    url = "/runtime/44/list/186/Solicitação de Pagamento"
    financeiro.find_element_by_xpath('//a[@href="'+url+'"]').click()
    
    #element = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]")
    #JavaScript
    #financeiro.execute_script("arguments[0].click();", element)

    #Filtro
    dataA = date.today().strftime("%d/%m/%Y")
    sleep(4)
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[2]/span[1]/span[1]/input").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[1]").click()
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[7]/div/div/div/div/div/div[1]/div/input").send_keys(dataA)
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
    #abrir a linha
    #financeiro.find_element_by_xpath('/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]').send_keys("\n")
    
    
    #abrir a linha (Laço para todos itens filtrados)
    
    sleep(3)
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]").click()
    #clique em certidoes
    #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
    #baixar
    #sleep(2)
    #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[1]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    #nota
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
    #baixar
    #sleep(2)
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    sleep(2)
    #clique em dados
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click()



    #imprimir
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[8]/div[2]/div/div/button").send_keys("\n")
    #baixar a capa
    #element = financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div")
    #JavaScript
    #financeiro.execute_script("arguments[0].click();", element)
    
    print('a')

    #financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]").click()
    #fechar
    financeiro.switch_to_frame(0)
    print(financeiro)
    print("aa")
    financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div").click()
    financeiro.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]").click()



'''
    #tramitar
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[3]").send_keys("\n")
    #processado
    financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[3]").click()
    #salvar e fechar
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click() #send_keys("\n")
    #refresh
    sleep(1)
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[1]").click() #send_keys("\n")
    '''
'''
def tramitar(baixar):    
    #baixando as certidoes
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[1]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[2]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[3]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr[1]/td[2]/div[2]/div/a").click()
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[4]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    #baixando as NF's
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    #imprimindo
    baixar.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[3]").click()
    #tramitando
    baixar.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[4]/div").click()
   ''' 




'''
page = financeiro.current_window_handle
for handle in financeiro.window_handles: 
        if handle != page: 
            lpage = handle 
    financeiro.switch_to.window(lpage)
'''
#financeiro.execute_script("document.getElementByXpath('/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[2]')[0].click()")