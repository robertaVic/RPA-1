# from selenium import webdriver
# from selenium.webdriver import Chrome
# from gerenciadorPlanilhas import tramitar_para_pago
from openpyxl import load_workbook
from time import sleep
import gerenciadorPastas
from gerenciadorPlanilhas import preencher_solicitacao_pagamento_avulso

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha


todosOsIds = sh1["B"]
tipo_avulso = sh1["A"]
statusRobo = sh1["R"]
statusFinanceiro = sh1["X"]
todos = list(range(7, len(tipo_avulso)))
listaId = []
listaLinha = []
listaStatusDiferente = []

novo = int(input("digite o numero: "))
status = "PROCESSADA"
ultima_linha = sh1.max_row
print(ultima_linha)
#if id ja existe nao adiciona
#ATUALIZAR
for i in todos:
    avulso = (tipo_avulso[i])
    if avulso.value == "SPA":
        listaId.append(todosOsIds[i].value)
        listaLinha.append(avulso.row)
        if statusRobo[i].value != statusFinanceiro[i].value:
            #print("diferente")
            listaStatusDiferente.append(statusRobo[i].row)

print(listaId)
print(listaLinha)    

for idd in range(len(listaId)):
    print(f"ID: {listaId[idd]} LINHA: {listaLinha[idd]}")
    if novo in listaId:
        #sobrescrever
        linhaa = listaId.index(novo)
        sh1.cell(row=listaLinha[linhaa], column=2, value=novo)
        wb.save(arquivo_excel)
    else:
        #adicionar um novo
        sh1.cell(row=ultima_linha+1, column=2, value=novo)
        # print("adicionando um novo")
        wb.save(arquivo_excel)
print(listaStatusDiferente)
for i in listaStatusDiferente:
    print(sh1[f"X{i}"].value)

    
            

# print(f"{novo} está na linha {listaLinha[linhaa]}")

    # for x in range(7, len(tipo_avulso)):
#         tipo = tipo_avulso[x]
#         if tipo.value == tipo_solicitacao: #== tipo_solicitacao
#             linhas_tramitacao.append(tipo.row)
    # if identificacao == int(dados_formulario[1]):
        # print("não pode ser adicionado")
    # else:    
    #     for coluna in range(len(dados_formulario)):
    #         sh1.cell(row=ultima_linha+1, column=coluna+1, value=dados_formulario[coluna])
    #     wb.save(arquivo_excel)
    
# driver = Chrome()
# #parte do sgp
# funcoes.chamarDriver(driver)
# funcoes.fazerLogin(driver)
# funcoes.encontrar_elemento_por_repeticao(driver,"/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div","link","SRB1",0.2)
# driver.get("https://tpf.madrix.app/runtime/44/list/190/Solicitação de Pgto Avulso")
# tramitar_para_pago("SPA", driver)