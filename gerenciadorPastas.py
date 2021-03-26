import os
from pathlib import Path
# from pathlib import Path
from datetime import date,datetime



today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")



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

def recuperar_diretorio_usuario():
    home = str(Path.home())
    return home
    
pref = recuperar_diretorio_usuario() + "\\OneDrive - tpfe.com.br\\RPA-DEV\\" 

