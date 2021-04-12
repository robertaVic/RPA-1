import os
from pathlib import Path
# from pathlib import Path
from datetime import date,datetime
from os import listdir
from os.path import isfile, join


today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")


def listar_arquivos_em_diretorios(diretorio):
    listaDeArquivos = [f for f in listdir(diretorio) if isfile(join(diretorio, f))]
    return listaDeArquivos


def criarPastaData(caminho, nomePasta):
    try:
        os.mkdir(caminho + nomePasta)
    except:
        print("Já existe a pasta")

def criarPastasFilhas(identificador):
    try:
        os.mkdir(caminho_da_pasta + data_em_texto + "\\" + identificador)
    except:
        print("Já existe a pasta")

def recuperar_diretorio_usuario():
    home = str(Path.home())
    return home

caminho_da_pasta = recuperar_diretorio_usuario() + "\\OneDrive - tpfe.com.br\\RPA-DEV\\" 

