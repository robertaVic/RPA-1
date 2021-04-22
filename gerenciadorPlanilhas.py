from openpyxl import load_workbook
import gerenciadorPastas

arquivo_excel = gerenciadorPastas.recuperar_diretorio_usuario() +"\\tpfe.com.br\\SGP e SGC - RPA\\Resultados\\Planilha de Acompanhamento de Solicitações Financeiras 2021.xlsx"
wb = load_workbook(arquivo_excel) #carregar o arquivo
sh1 = wb.worksheets[0] #carregar a primeira planilha

#retorna a ultima linha da planilha


def preencher_solicitacao_pagamento_avulso(dados_formulario):
    ultima_linha = sh1.max_row
    for coluna in range(len(dados_formulario)):
        sh1.cell(row=ultima_linha+1, column=coluna+1, value=dados_formulario[coluna])
    wb.save(arquivo_excel)



