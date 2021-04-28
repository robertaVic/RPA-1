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

def tramitar_para_pago(tipo_de_solicitacao, dado):
    for i in listaStatusDiferente:
        statusR = statusRobo[i].value
        print(statusR)
        if statusR == "Processada":
            identificacao = todosOsIds[i].value
            print(identificacao)
            valor = sh1[f"L{i}"].value
            data = (sh1[f"Y{i}"].value).strftime("%d/%m/%Y")
            print(f"Valor pago: {valor} data: {data}")
            status = sh1[f"X{i}"].value
            if dado == "id":
                return identificacao
            elif dado == "valor":
                return valor
            elif dado == "data":
                return data     
            elif dado == "opcao":       
                if status == "PAGO":   
                    return "1"  
                    sh1.cell(row=i, column=18, value="PAGO")
                    wb.save(arquivo_excel)
                elif status ==  "PARCIALMENTE PAGO":
                    return "2"
                    sh1.cell(row=i, column=18, value="PARCIALMENTE PAGO")
                    wb.save(arquivo_excel)


        


