# from selenium import webdriver
# from selenium.webdriver import Chrome
# from gerenciadorPlanilhas import tramitar_para_pago
from openpyxl import load_workbook
from time import sleep
import gerenciadorPastas

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

ultima_linha = sh1.max_row
todosOsIds = sh1["B"]
tipo = sh1["A"]
statusRobo = sh1["R"]
statusFinanceiro = sh1["X"]
todos = list(range(7, ultima_linha))
listaId = []
listaLinha = []
listaStatusDiferente = []
dados = []

def ler_dados_da_planilha(tipo_de_solicitacao):
    for i in todos:
        avulso = tipo[i]
        if avulso.value == tipo_de_solicitacao:
            listaId.append(todosOsIds[i].value)
            listaLinha.append(avulso.row)
            statusRo = statusRobo[i]
            statusFinan = statusFinanceiro[i]
            if statusRo.value != statusFinan.value:
                #print("diferente")
                cadaSolicitacao = []
                cadaSolicitacao.append(todosOsIds[statusRo.row].value)
                cadaSolicitacao.append(sh1[f"L{statusRo.row}"].value)
                cadaSolicitacao.append((sh1[f"Y{statusRo.row}"].value).strftime("%d/%m/%Y"))
                cadaSolicitacao.append(sh1[f"X{statusRo.row}"].value)
                listaStatusDiferente.append(statusRo.row)
                dados.append(cadaSolicitacao)
    return dados     
    print(listaId)
    print(listaStatusDiferente)    
    print(dados)
def quantidade_para_tramitacao():
    return listaStatusDiferente        
print(ler_dados_da_planilha("SPA"))    

def tramitar_para_pago(tipo_de_solicitacao, dado, iteracao):
    ler_dados_da_planilha("SPA")
    if dado == "id":
        return (dados[iteracao])[0]
    elif dado == "valor":
        return (dados[iteracao])[1]
    elif dado == "data":
        return (dados[iteracao])[2]     
    elif dado == "opcao":       
        if (dados[iteracao])[3] == "PAGO":     
            sh1.cell(row=listaStatusDiferente[iteracao], column=18, value="PAGO")
            wb.save(arquivo_excel)
            return "1"
        elif (dados[iteracao])[3] ==  "PARCIALMENTE PAGO":
            sh1.cell(row=listaStatusDiferente[iteracao], column=18, value="PARCIALMENTE PAGO")
            wb.save(arquivo_excel) 
            return "2"



for i in range(len(quantidade_para_tramitacao())):
    tramitar_para_pago("SPA", "id", i)
    tramitar_para_pago("SPA", "valor", i)
    tramitar_para_pago("SPA", "data", i)
    tramitar_para_pago("SPA", "opcao", i)           
# novo = int(input("digite o numero: "))
# status = "PROCESSADA"
# print(ultima_linha)
# dados_formulario = ["SPA", "0890034", "0899779090", "CERRADO", "", "", "", "", "", "", "DEATAJSDI", "HKJIUYDSFI", "LJHZKJFL"]
# #if id ja existe nao adiciona
# #ATUALIZAR
# def preencher_solicitacao(dados_formulario, tipo_de_solicitacao):
#     for i in todos:
#         avulso = (tipo_avulso[i])
#         if avulso.value == tipo_de_solicitacao:
#             listaId.append(todosOsIds[i].value)
#             listaLinha.append(avulso.row)
#             statusRo = statusRobo[i]
#             statusFinan = statusFinanceiro[i]
#             if statusRo.value != statusFinan.value:
#                 #print("diferente")
#                 listaStatusDiferente.append(statusRobo[i].row)
#     print(listaId)            
#     for coluna in range(len(dados_formulario)):
#         #print(f"ID: {listaId[idd]} LINHA: {listaLinha[idd]}")
#         if dados_formulario[1] in listaId:
#             print("JA EXISTE, SOBRESCREVER")
#             #sobrescrever
#             linhaa = listaId.index(dados_formulario[1])
#             sh1.cell(row=listaLinha[linhaa], column=coluna+1, value=dados_formulario[coluna])
#             wb.save(arquivo_excel)
#         else:
#             print("NÃO EXISTE, SALVAR NOVO")
#             #adicionar um novo
#             sh1.cell(row=ultima_linha+1, column=coluna+1, value=dados_formulario[coluna])
#             # print("adicionando um novo")
#             wb.save(arquivo_excel)
#         print("SALVOU")  
#     print(ultima_linha) 


# preencher_solicitacao(dados_formulario, "SPA")
    
            

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