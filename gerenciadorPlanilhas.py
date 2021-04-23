from openpyxl import load_workbook
import gerenciadorPastas
import funcoes
from time import sleep

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

#retorna a ultima linha da planilha
todosOsIds = sh1["B"]
tipo_avulso = sh1["A"]
statusRobo = sh1["R"]
todos = list(range(7, len(todosOsIds)))

def preencher_solicitacao_pagamento_avulso(dados_formulario):
    ultima_linha = sh1.max_row
    for coluna in range(len(dados_formulario)):
        sh1.cell(row=ultima_linha+1, column=coluna+1, value=dados_formulario[coluna])
    wb.save(arquivo_excel)

def tramitar_para_pago(tipo_solicitacao, driver):
    linhas_tramitacao = []

    for x in range(7, len(tipo_avulso)):
        tipo = tipo_avulso[x]
        if tipo.value == "SPA": #== tipo_solicitacao
            linhas_tramitacao.append(tipo.row)
    print(linhas_tramitacao)

    for linha in linhas_tramitacao:
        statusR = statusRobo[linha].value
        print(statusR)
        if statusR == "Processada":
            identificacao = todosOsIds[linha].value
            print(identificacao)
            valor = sh1[f"L{linha}"].value
            data = (sh1[f"Y{linha}"].value).strftime("%d/%m/%Y")
            print(f"Valor pago: {valor} data: {data}")
            status = sh1[f"X{linha}"].value
            #parte do sgp
            funcoes.chamarDriver(driver)
            funcoes.fazerLogin(driver)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","link","SRB1",0.2)
            driver.get("https://tpf.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro",0.2)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[6]","click","filtro",0.2)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro",0.2)
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(identificacao)
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
            funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","link","SRB1",0.2)
            driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()
            driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]").click()
            ##------------------------------##
            if status == "PAGO":
                funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","link","SRB1",0.2)     
                print("PAGO NO SGP")
                sh1.cell(row=linha, column=18, value="PAGO")
                wb.save(arquivo_excel)
            elif status ==  "PARCIALMENTE PAGO":
                funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","link","SRB1",0.2)     
                print("PARCIALMENTE PAGO NO SGP")
                sh1.cell(row=linha, column=18, value="PARCIALMENTE PAGO")
                wb.save(arquivo_excel)
        


