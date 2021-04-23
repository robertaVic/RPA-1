from openpyxl import load_workbook
import gerenciadorPastas
from datetime import date, datetime
from selenium import webdriver
from selenium.webdriver import Chrome
from time import sleep
from selenium.webdriver.common.action_chains import ActionChains 

driver = Chrome()
builder = ActionChains(driver)

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021 .xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

todosOsIds = sh1["B"]
statusRobo = sh1["R"]
#TOP

todos = list(range(7, len(todosOsIds)))    

driver.maximize_window()
driver.get("https://tpf.madrix.app/")
driver.implicitly_wait(10)
driver.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[1]/div/div/div/input").send_keys("roberta.costa")
driver.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[2]/div/div/div/input").send_keys("123")
driver.find_element_by_xpath("/html/body/div/div/div[2]/main/div[2]/div/div/div/section/form/div[3]/div/button/span[1]").click()
driver.implicitly_wait(40)
driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div").click()
driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div[1]").click()
url = "/runtime/44/list/190/Solicitação de Pgto Avulso"
#clicar na parte de pagamento avulso
driver.find_element_by_xpath('//a[@href="'+url+'"]').click()
sleep(10)

for i in todos:
    statusR = statusRobo[i].value
    print(statusR)
    if statusR == "Processada":
        idd = todosOsIds[i].value
        print(idd)
        valor = sh1[f"L{i+1}"].value
        data = (sh1[f"Y{i+1}"].value).strftime("%d/%m/%Y")
        print(f"Valor pago: {valor}  data: {data}")
        linha = (todosOsIds[i].row)
        status = sh1[f"X{linha}"].value
        sleep(5)
        filtro = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")
        try:
            filtro_click = builder.click(filtro)
            filtro_click.perform()
        except:
            print("nao clicou")
        try:
            filtro_click = builder.click(filtro)
            filtro_click.perform()
        except:
            print("nao clicou") 
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[1]/div[1]/button").click()
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(idd)
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
        driver.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[7]").send_keys("\n")
        driver.find_element_by_xpath("/html/body/div[6]/div[1]").click()
        print(20*"=")
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
        sleep(5)
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input").click()
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
        sleep(3)
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]").send_keys("\n")
        driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(data)
        driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(valor)
        try:
            driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]").send_keys("\n")
        except:
            pass
        try:
            driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]").send_keys("\n")
        except:
            pass       
        driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()
        driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]").click()
        if status == "PAGO":
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/ul/div[1]").click()
            print("PAGO NO SGP")
            sh1.cell(row=i+1, column=18, value="PAGO")
            wb.save(arquivo_excel)
        elif status == "PARCIALMENTE PAGO":
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/ul/div[2]").click()
            print("PARCIALMENTE PAGO NO SGP")
            sh1.cell(row=i+1, column=18, value="PARCIALMENTE PAGO")
            wb.save(arquivo_excel)

       


    
    # print(dataString)





# for row in sh1.rows:
#     for cell in row:
#         print(cell.value)        