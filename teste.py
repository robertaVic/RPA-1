
# from openpyxl import load_workbook
# arquivo_excel = load_workbook(filename="C:\\Users\\Usuario\\Downloads\\Solicitação de Pgto Avulso.xlsx")
# arquivo_excel.sheetnames
# for d in arquivo_excel['Sheet1'].iter_rows(values_only = True):
#     print(d)
import pandas as pd

df = pd.read_excel("C:\\Users\\Usuario\\Downloads\\Solicitação de Pgto Avulso.xlsx")
# print(df.values)
planilha = df.values

linhaExcel = []
conjunto = []

for linha in planilha:
    print("-------")
    # print(linha)
    data = linha[8]
    valor = linha[9]
    if data != 'Nat' and valor != 'nan':
        idSolicitacao = linha[0]
        print(idSolicitacao)
    print(data)
    print(valor)
    conjunto.append(linha)
    for celula in linha:
        if not celula:
            print(celula)
        # print(celula)
        # linhas.append(celula)

# conjunto.append(linhaExcel)    

# print(conjunto)
        








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