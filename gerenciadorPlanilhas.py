from openpyxl import load_workbook
import gerenciadorPastas
import funcoes
from time import sleep
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
from openpyxl.cell import Cell

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha


def ler_dados_da_planilha(tipo_de_solicitacao):
    ultima_linha = sh1.max_row
    dados = []
    listaId = []
    todos = list(range(8, ultima_linha))
    for i in todos:
        tipo = sh1['A']
        todosOsIds = sh1["B"]
        statusRobo = sh1["R"]
        statusFinanceiro = sh1["X"]
        # avulso = tipo[i]
        if tipo[i].value == tipo_de_solicitacao:
            listaId.append(todosOsIds[i].value)
            statusRo = statusRobo[i]
            statusFinan = statusFinanceiro[i]
            if statusRo.value != statusFinan.value:
                cadaSolicitacao = []
                status = statusRo.row
                #geral = [todosOsIds[status].value, sh1[f"L{status}"].value, (sh1[f"Y{status}"].value).strftime("%d/%m/%Y"), sh1[f"X{status}"].value, status]
                cadaSolicitacao.append(sh1[f"B{status}"].value)
                cadaSolicitacao.append(sh1[f"L{status}"].value)
                cadaSolicitacao.append((sh1[f"Y{status}"].value).strftime("%d/%m/%Y"))
                cadaSolicitacao.append(sh1[f"X{status}"].value)
                cadaSolicitacao.append(status)
                dados.append(cadaSolicitacao)
    return dados     

def preencher_solicitacao_na_planilha(dados_formulario, tipo_de_solicitacao):
    ultima_linha = sh1.max_row
    todos = list(range(8, ultima_linha))
    listaId = []
    for i in todos:
        todosOsIds = sh1["B"]
        tipo = sh1['A']
        # avulso = tipo[i]
        if tipo[i].value == tipo_de_solicitacao:
            listaId.append(todosOsIds[i].value)

    for coluna in range(len(dados_formulario)):
        #print(f"ID: {listaId[idd]} LINHA: {listaLinha[idd]}")
        if dados_formulario[1] in listaId:
            print("JA EXISTE, SOBRESCREVER")
            #sobrescrever
            linhaa = listaId.index(dados_formulario[1])
            sh1.cell(row=listaLinha[linhaa], column=coluna+1, value=dados_formulario[coluna])
            wb.save(arquivo_excel)
        else:
            print("NÃO EXISTE, SALVAR NOVO")
            #adicionar um novo
            sh1.cell(row=ultima_linha+1, column=coluna+1, value=dados_formulario[coluna])
            # print("adicionando um novo")
            wb.save(arquivo_excel)
        print("SALVOU")  
    print(ultima_linha)

def atualizar_status_na_planilha(linha):
    #pegar o status daqui mesmo
    myFill = PatternFill(start_color='A9D08E', 
                    end_color='A9D08E', 
                    fill_type = 'solid')
    sh1.cell(row=linha, column=18, value= sh1[f"X{linha}"].value).fill = myFill
    wb.save(arquivo_excel)     

# def quantidade_para_tramitacao():
#     return len(ler_dados_da_planilha())        

# def tramitar_para_pago(tipo_de_solicitacao, dado, iteracao): #atualizar status
#     ler_dados_da_planilha(tipo_de_solicitacao)
#     if dado == "id":
#         return (dados[iteracao])[0]
#     elif dado == "valor":
#         return (dados[iteracao])[1]
#     elif dado == "data":
#         return (dados[iteracao])[2]     
#     elif dado == "opcao":       
#         if (dados[iteracao])[3] == "PAGO":     
#             sh1.cell(row=listaStatusDiferente[iteracao], column=18, value="PAGO")
#             wb.save(arquivo_excel)
#             return "1"
#         elif (dados[iteracao])[3] == "PARCIALMENTE PAGO":
#             sh1.cell(row=listaStatusDiferente[iteracao], column=18, value="PARCIALMENTE PAGO")
#             wb.save(arquivo_excel) 
#             return "2"





