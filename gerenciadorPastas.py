import os
from datetime import date,datetime

today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")

pref = "C:\\Users\\Usuario\\OneDrive - tpfe.com.br\\" 

def criarPastaData(pref, nomePasta):
    try:
        os.mkdir(pref + nomePasta)
    except:
        print("Já existe a pasta")
def criarPastasFilhas(identificador):
    try:
        os.mkdir(pref + data_em_texto + "\\" + identificador)
    except:
        print("Já existe a pasta")

criarPastaData(pref, data_em_texto)

