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

def criarPastasFilhas(tipo_solicitacao,identificador):
    print(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + tipo_solicitacao +"\\"+ data_em_texto + "\\" + identificador)
    #verificar os erros na hora de criar a pasta, pois ela diz que ja existe, sem ter
    try:
        os.mkdir(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + tipo_solicitacao +"\\"+ data_em_texto + "\\" + identificador)
    except:
        print("Já existe a pasta")
    try:
        os.mkdir(recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" + tipo_solicitacao +"\\"+ data_em_texto + "\\" + identificador)
    except:
        print("Já existe a pasta")

def recuperar_diretorio_usuario():
    home = str(Path.home())
    return home

caminho_da_pasta = recuperar_diretorio_usuario() + "\\tpfe.com.br\\SGP e SGC - RPA\\" 

def remover_arquivos_da_raiz(diretorio):
    arquivos = listar_arquivos_em_diretorios(diretorio)
    for arquivo in arquivos:
        try:
            os.remove(diretorio+arquivo)
        except:
            print("Não deletou o arquivo")   
