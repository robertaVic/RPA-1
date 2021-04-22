from openpyxl import load_workbook
import gerenciadorPastas
from datetime import date, datetime

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021 .xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

status = sh1["X"]
linhasTramitacao = []

# Print the contents
for x in range(7, len(status)):
    celula = (status[x].value)
    if celula == "PAGO":
        print(celula)
        print("PAGO NO SGP")
        linha = (status[x].row)
        linhasTramitacao.append(linha)

print(linhasTramitacao) 

for x in linhasTramitacao:
    valor = sh1[f"L{x}"].value
    data = (sh1[f"Y{x}"].value).strftime("%d/%m/%Y")
    print(f"Valor pago: {valor}, data: {data}")
    
    # print(dataString)
        
# for row in sh1.rows:
#     for cell in row:
#         print(cell.value)        