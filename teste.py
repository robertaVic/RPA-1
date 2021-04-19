from datetime import date

data_solicitacao = date(2021, 4, 27)
data_solicitacao.strftime("%d.%m.%Y")
today = date.today()
#formataçao da data para o modelo de pasta do financeiro
data_em_texto = today.strftime("%d.%m.%Y")

dias = list(range(1,16))
print(dias)

# if data_solicitacao.day in dias:
#     print("pagar até dia 30")
# else:
#     print("pagar até dia 04 do outro mês")    
    
if today.day in dias:
    print("puxar as solicitações de 1 a 15")
    data_inicio = date(today.year, today.month, 1)
    data_fim = date(today.year, today.month, 15)  
    data_fim_formatada = data_fim.strftime("%d/%m/%Y")
    data_inicio_formatada = data_inicio.strftime("%d/%m/%Y")
else:
    print("puxar as solicitações de 16 a 30")
    data_inicio = date(today.year, today.month, 16)
    data_fim = date(today.year, today.month, 30)  
    data_fim_formatada = data_fim.strftime("%d/%m/%Y")
    data_inicio_formatada = data_inicio.strftime("%d/%m/%Y")

print(f"inicio {data_inicio_formatada}\nData de fim {data_fim_formatada}")      