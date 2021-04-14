
from openpyxl import load_workbook
arquivo_excel = "C:\\Users\\Usuario\\OneDrive - tpfe.com.br\\RPA-DEV\\Planilha de Acompanhamento de Solicitações Financeiras 2021 (1).xlsx"
wb = load_workbook(arquivo_excel)
sh1 = wb.worksheets[0]

identificaçao = "986690"
razaoSocial = "CELL AUTO"
valor = "tal"
forma = "dinheiro"
cnpj = "98790-8989"
banco = "BRA"
conta = "77876-a"
agencia = "fdf"
tipo = "nenhum"
natureza = "nenhuma"
valorPago = "x"
dataSolicitada = "34/05/2021"

maximo = sh1.max_row
print(maximo) 

# for row in sh1:
#     if not row:      
#         pro = row.index
#         print(pro) 

   

# minimo = lista[0]
lista = list(range(8,maximo))
solicitaçao = list(range(5))
for x in solicitaçao:
    sh1[f"A{lista[0]}"].value = identificaçao
    sh1[f"B{lista[0]}"].value = cnpj
    sh1[f"C{lista[0]}"].value = razaoSocial
    sh1[f"D{lista[0]}"].value = forma
    sh1[f"E{lista[0]}"].value = banco
    sh1[f"F{lista[0]}"].value = agencia
    sh1[f"G{lista[0]}"].value = conta
    sh1[f"H{lista[0]}"].value = tipo
    sh1[f"I{lista[0]}"].value = natureza
    sh1[f"J{lista[0]}"].value = valor
    sh1[f"K{lista[0]}"].value = valorPago
    sh1[f"L{lista[0]}"].value = dataSolicitada
    lista.pop(0)
print(lista) 
solicita = list(range(10))   
for x in solicita:
    sh1[f"A{lista[0]}"].value = identificaçao
    sh1[f"B{lista[0]}"].value = cnpj
    sh1[f"C{lista[0]}"].value = razaoSocial
    sh1[f"D{lista[0]}"].value = forma
    sh1[f"E{lista[0]}"].value = banco
    sh1[f"F{lista[0]}"].value = agencia
    sh1[f"G{lista[0]}"].value = conta
    sh1[f"H{lista[0]}"].value = tipo
    sh1[f"I{lista[0]}"].value = natureza
    sh1[f"J{lista[0]}"].value = valor
    sh1[f"K{lista[0]}"].value = valorPago
    sh1[f"L{lista[0]}"].value = dataSolicitada
    lista.pop(0)
                                                               


print(lista)

wb.save(arquivo_excel)




# for cell in sh1["A"]:
#     if cell.value is not None:
#         print(cell.row)
#         break
# else:
#     print(cell.row + 1)
# while len(lista) > 0:
#     print(f"célula {lista[0]}")
#     lista.pop(0) 
#     print(lista)
# for i in range(sh1.nrows) :                                     
#     for j in range (sh1.ncols) :                                
#         ptrow=i                                                   
#         if(sh1.cell_value(ptrow,j)=="") :                       
#             count +=1                                             
#         if (count==sh1.ncols):                                  
#             return ptrow                                          
#         else:                                                     
#             continue                                              
                               
# rownum=rtrow()                                                             
# rownum=rownum+1                                                                 
# print(f"The presence of an empty row is at :{rownum}")
# linha = sh1.max_row
# print(linha)
    


    






# import pandas as pd

# df = pd.read_excel("C:\\Users\\Usuario\\Downloads\\Solicitação de Pgto Avulso.xlsx")
# # print(df.values)
# planilha = df.values

# linhaExcel = []
# conjunto = []

# for linha in planilha:
#     print("-------")
#     # print(linha)
#     data = linha[8]
#     valor = linha[9]
#     if data != 'Nat' and valor != 'nan':
#         idSolicitacao = linha[0]
#         print(idSolicitacao)
#     print(data)
#     print(valor)
#     conjunto.append(linha)
#     for celula in linha:
#         if not celula:
#             print(celula)









# print(arquivo_excel)
# c1 = arquivo_excel['C1']
# max_linha = arquivo_excel.max_row
# max_coluna = arquivo_excel.max_column
# for i in range(1, max_linha + 1):
#     for j in range(1, max_coluna + 1):
#         print(arquivo_excel.cell(row=i, column=j).value, end=" - ")

# book = xlrd.open_workbook("")
# print ("Número de abas: ", book.nsheets)
# print ("Nomes das Planilhas:", book.sheet_names())
# sh = book.sheet_by_index(0)
# print(sh.name, sh.nrows, sh.ncols)
# print("Valor da célula C6 é ", sh.cell_value(rowx=5, colx=2))
# for rx in range(sh.nrows):
#     print(sh.row(rx))