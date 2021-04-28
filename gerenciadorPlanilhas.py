from openpyxl import load_workbook
import gerenciadorPastas
import funcoes
from time import sleep

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

#retorna a ultima linha da planilha 
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


def preencher_solicitacao_na_planilha(dados_formulario, tipo_de_solicitacao):
    for i in todos:
        avulso = (tipo[i])
        if avulso.value == tipo_de_solicitacao:
            listaId.append(todosOsIds[i].value)
            listaLinha.append(avulso.row)
            statusRo = statusRobo[i]
            statusFinan = statusFinanceiro[i]
            if statusRo.value != statusFinan.value:
                #print("diferente")
                listaStatusDiferente.append(statusRobo[i].row)
    print(listaId)            
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

def quantidade_para_tramitacao():
    return listaStatusDiferente        

def tramitar_para_pago(tipo_de_solicitacao, dado, iteracao):
    ler_dados_da_planilha(tipo_de_solicitacao)
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
        elif (dados[iteracao])[3] == "PARCIALMENTE PAGO":
            sh1.cell(row=listaStatusDiferente[iteracao], column=18, value="PARCIALMENTE PAGO")
            wb.save(arquivo_excel) 
            return "2"


        


