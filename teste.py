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
# from gerenciadorPlanilhas import ler_dados_da_planilha, atualizar_status_na_planilha, preencher_solicitacao_na_planilha



# print(arquivos)
def listar_arquivos_em_diretorios(diretorio):
    listaDeArquivos = [f for f in os.listdir(diretorio) if isfile(join(diretorio, f))]
    return listaDeArquivos
arquivos = listar_arquivos_em_diretorios(gerenciadorPastas.recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA")
#criando um laço para mover cada um para sua pasta especifica
for arquivo in arquivos:
    print(arquivo)
    fileExt = r".crdownload"
    if _.endswith(fileExt):
        print("ok")




# data_em_texto = date.today().strftime("%d/%m/%Y")
# arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
# wb = load_workbook(arquivo_excel) #carregar o arquivo
# sh1 = wb.worksheets[0]


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
    
            
