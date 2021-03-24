import os
import getpass
# from pathlib import Path
from datetime import date,datetime

usuario = getpass.getuser()

today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")

pref = "C:\\Users\\"+ usuario +"\\OneDrive - tpfe.com.br\\RPA-DEV" 

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



