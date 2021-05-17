
from gerenciadorPastas import listar_arquivos_em_diretorios
import os


def remover_arquivos_da_raiz(diretorio):
    arquivos = listar_arquivos_em_diretorios(diretorio)
    for arquivo in arquivos:
        try:
            os.remove("C:\\Users\\Carlos Ivan\\Desktop\\Testes\\" + arquivo)
        except:
            print("Não deletou o arquivo")
#C:\Users\Carlos Ivan\Desktop\Testes
remover_arquivos_da_raiz("C:\\Users\\Carlos Ivan\\Desktop\\Testes")
#remover_arquivos_da_raiz("C:\\Users\\Carlos Ivan\\tpfe.com.br\\SGP e SGC - RPA\\Testes")