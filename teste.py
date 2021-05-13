# from selenium import webdriver
# from selenium.webdriver import Chrome
# from gerenciadorPlanilhas import tramitar_para_pago
import os
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell
from datetime import date
from openpyxl import load_workbook
from time import sleep
import gerenciadorPastas
from os.path import isfile, join
# from gerenciadorPlanilhas import ler_dados_da_planilha, atualizar_status_na_planilha, preencher_solicitacao_na_planilha

os.remove(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\.849C9593-D756-4E56-8D6E-42412F2A707B")


        




data_em_texto = date.today().strftime("%d/%m/%Y")
arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021R.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0]

# rows2 = ["beta", "betinha", "betona"]
# ultima_linha = (sh1.max_row)
# print(ultima_linha)

# for row in rows2:
#     maximo_tentativas = 0
#     while maximo_tentativas < 40: 
#         if len(rows2) > 0:
#             print(row)
#             sleep(3)
#             maximo_tentativas = 40
#         else:
#             maximo_tentativas += 1
# ler_dados_da_planilha("SPA") 
# print(len(ler_dados_da_planilha("SPA"))) 

# print((ler_dados_da_planilha("SPA")))
# for i in range(len(ler_dados_da_planilha("SPA"))):
#     print(ler_dados_da_planilha("SPA")[0][0])
#     ler_dados_da_planilha("SPA").pop(-1)
#     print((ler_dados_da_planilha("SPA")))
#     if ler_dados_da_planilha("SPA")[i][2] == data_em_texto:
#         print("Pode pagar hoje")
#         myFill = PatternFill(start_color='FF3399', 
#                     end_color='FF3399', 
#                     fill_type = 'solid')

#         sh1.cell(row=ler_dados_da_planilha("SPA")[i][4],column=18).fill = myFill
#         wb.save(arquivo_excel)
    #print(dados[3])
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
    
            
