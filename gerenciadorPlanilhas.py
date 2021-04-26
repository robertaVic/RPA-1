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
statusFinanceiro = sh1["X"]
todos = list(range(7, len(tipo_avulso)))
listaId = []
listaLinha = []
listaStatusDiferente = []

for i in todos:
    avulso = (tipo_avulso[i])
    if avulso.value == "SPA":
        listaId.append(todosOsIds[i].value)
        listaLinha.append(avulso.row)
        if statusRobo[i].value != statusFinanceiro[i].value:
            #print("diferente")
            listaStatusDiferente.append(statusRobo[i].row)

def preencher_solicitacao_pagamento_avulso(dados_formulario):
    for idd in range(len(listaId)):
        print(f"ID: {listaId[idd]} LINHA: {listaLinha[idd]}")
        if dados_formulario[1] in listaId:
            #sobrescrever
            linhaa = listaId.index(dados_formulario[1])
            sh1.cell(row=listaLinha[linhaa], column=idd+1, value=dados_formulario[1])
            wb.save(arquivo_excel)
        else:
            #adicionar um novo
            sh1.cell(row=ultima_linha+1, column=idd+1, value=dados_formulario[1])
            # print("adicionando um novo")
            wb.save(arquivo_excel)

# def tramitar_para_pago(tipo_solicitacao, driver):
#     linhas_tramitacao = []
#     for x in range(7, len(tipo_avulso)):
#         tipo = tipo_avulso[x]
#         if tipo.value == tipo_solicitacao: #== tipo_solicitacao
#             linhas_tramitacao.append(tipo.row)
#     print(linhas_tramitacao)
#     for linha in linhas_tramitacao:
#         statusR = statusRobo[linha].value
#         print(statusR)
#         if statusR == "Processada":
#             identificacao = todosOsIds[linha].value
#             print(identificacao)
#             valor = sh1[f"L{linha}"].value
#             data = (sh1[f"Y{linha}"].value).strftime("%d/%m/%Y")
#             print(f"Valor pago: {valor} data: {data}")
#             status = sh1[f"X{linha}"].value
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div/div/div","click","filtro",0.2)
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/ul/li[6]","click","filtro",0.2)
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]","click","filtro",0.2)
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[1]/div[1]/button","click","filtro",0.2)
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[1]/div/div/div/div/input").send_keys(identificacao)
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
#             driver.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[7]").send_keys("\n")
#             driver.find_element_by_xpath("/html/body/div[6]/div[1]").click()
#             print(20*"=")
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
#             sleep(5)
#             driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr/td[2]/span/span[1]/input").click()
#             driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[3]").click()
#             sleep(3)
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[4]/div/div/div/div[1]/div[1]/div[1]/div/div/span/div/button[1]","click","SRB1",0.2)
#             driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[1]/div/div/div/input").send_keys(data)
#             driver.find_element_by_xpath("/html/body/div[8]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div/div[1]/div[2]/div/div/div/input").send_keys(str(valor))
#             funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[8]/div[3]/div/div/div/div[4]/fieldset/button[2]","click","SRB1",0.2)
#             funcoes.encontrar_elemento_por_repeticao(driver, "//*[@id='main1']", "click", "SPA", 0.2)
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div/input").send_keys(str(valor)
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[2]/div/div[1]/div/div/div/input").send_keys(data)
#             funcoes.encontrar_elemento_por_repeticao(driver, "/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[2]", "click", "SPA", 0.2)
#             driver.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()
#             driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]").click()
#             ##------------------------------##
#             if status == "PAGO":
#                 funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[1]","click","SRB1",0.2)     
#                 print("PAGO NO SGP")
#                 sh1.cell(row=linha, column=18, value="PAGO")
#                 wb.save(arquivo_excel)
#             elif status ==  "PARCIALMENTE PAGO":
#                 funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[5]/div[3]/div/div[2]/ul/div[2]","click","SRB1",0.2)     
#                 print("PARCIALMENTE PAGO NO SGP")
#                 sh1.cell(row=linha, column=18, value="PARCIALMENTE PAGO")
#                 wb.save(arquivo_excel)
        


