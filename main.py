#Codigo Principal
from selenium import webdriver
from selenium.webdriver import Chrome
from funcoes import *
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from datetime import date,datetime
from gerenciadorPastas import criarPastaData, recuperar_diretorio_usuario
from pgAvulso import pagamentoAvulso

#carlos

#retornar a data atual
today = date.today()
#formataçao da data para o modelo de pasta do financeiro
data_em_texto = today.strftime("%d.%m.%Y")

#diretório padrão para todos os arquivos serem baixados nele
diretorio_padrao = "\\OneDrive - tpfe.com.br\\RPA-DEV\\" + data_em_texto + "\\"

#padrão do chrome 
chrome_options = padraoChrome(diretorio_padrao)

#caminho onde a pasta será criada
caminho_da_pasta = recuperar_diretorio_usuario() + "\\OneDrive - tpfe.com.br\\RPA-DEV\\" 

#inserindo as opçoes do chrome no driver
driver = Chrome(chrome_options=chrome_options)

#criar a pasta do dia no onedrive
criarPastaData(caminho_da_pasta, data_em_texto)

#Inicia o navegador
chamarDriver(driver)

#Login
fazerLogin(driver)

#entra no pagamento avulso
pagamentoAvulso(driver)


#volta para o sgp para baixar as NF'S
# baixarNf(driver)
