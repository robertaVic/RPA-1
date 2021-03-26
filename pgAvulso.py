from time import sleep
from datetime import date
import gerenciadorPastas
from funcoes import padraoChrome
import shutil
import getpass
import os

#retornar o usuario da maquina utilizada pelo robo
usuario = getpass.getuser()
#retornar a data atual
today = date.today()
#formataçao da data para o modelo de pasta do financeiro
data_em_texto = today.strftime("%d.%m.%Y")

def pagamentoAvulso(financeiro):
    financeiro.implicitly_wait(120)
    # sleep(10)
    # financeiro.refresh()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div[1]").click()
    url = "/runtime/44/list/190/Solicitação de Pgto Avulso"
    financeiro.find_element_by_xpath('//a[@href="'+url+'"]').click()

    #Filtro
    #dataA = date.today().strftime("%d/%m/%Y")
    sleep(10)
    #limpar filtro  => financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[2]/button").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[3]").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[1]").click()
    print(5*"*")
    # financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[10]/div/div/div/div/div/div[1]/div/input").send_keys(dataA)
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
    
    #abrir a linha (Laço para todos itens filtrados)
    sleep(10)
    tbody1 = financeiro.find_element_by_xpath("//*[@id='mainContent']/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody")
    rows1 = tbody1.find_elements_by_tag_name("tr")

    global solicitaçao
    solicitaçao = [] 
    for row in range(len(rows1)):
        solicitaçao.append(row) #armazenando os indices de cada linha

    print(solicitaçao)
    pastas = []
    for linha in solicitaçao: #para cada linha na solicitação
        index1 = str(linha+1)
        #para cada linha guardar as informações dela
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[4]/div").get_attribute("innerText")
        global razao
        #armazenando a razao social de cada solicitaçao
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[6]/div").get_attribute("innerText")
        #criando um modelo de nome de pastas para serem salvas(igualmente ao modelo do financeiro)
        nomeDaPasta = (f"ID {identificador} {razao}")
        print(nomeDaPasta)
        #adicionando cada modelo de nome de pasta para uma lista
        pastas.append(nomeDaPasta)
    for cadaPasta in pastas:
        #para cada solicitaçao, criar as pastas respectivas
        gerenciadorPastas.criarPastasFilhas(cadaPasta)
    #tempo para salvar todas pastas
    sleep(10)
    #poderiamos reutilizar o laço de "linha in solicitaçao", mas preciso que o nome da pasta seja um elemento iteravel
    #para podermos usar o nome da pasta para direcionar o caminho das nf's
    #para cada solicitaçao, clique nelas e baixe as nf e imprima
    for cadaPasta in pastas:
        #pega a posição dela na lista
        posiçao = pastas.index(cadaPasta)
        #acrescenta mais um para ser inserido corretamente no xpath
        index2 = str(posiçao+1)
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index2 +"]/td[2]/span/span[1]/input").click()
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()

        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        rows2 = tbody2.find_elements_by_tag_name("a") # pega todas as linhas que contem nf
        #criaçao de uma lista para armazenar os indices de cada nota fiscal
        nf = [] 
        #laço para adicionar os valores de *indices* na lista
        for row in range(len(rows2)):
            nf.append(row)
            print(row)
        print(nf)
        #nova lista criada apenas para guardar os *nomes* de cada nota fiscal(difere da lista anterior)
        nomeDaNf = []
        #para cada nf, baixar cada uma delas
        for nota in nf:
            index3 = str(nota+1)
            #capturando os nomes de cada nota fiscal
            nomeNf = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr["+ index3 +"]/td[2]/div[2]/div/a").get_attribute("innerText")
            financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr["+ index3 +"]/td[2]/div[2]/div/a").click()
            #imprime o nome de cada nf
            print(nomeNf)
            #guarda os nomes na lista nomeDaNf
            nomeDaNf.append(nomeNf)
            #imprime se foram baixadas
            print(f"{index3}º NF baixada!")
        sleep(10)    
        #imprimindo    
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click() 
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button").click()   
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr/td[1]").click()
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()

#após pegar os valores do formulário, ir no onedrive e criar pasta com o ID + RAZAO SOCIAL 
# def chamarSharepoint(nuvem, Keys):
#     nuvem.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')

#     nuvem.find_element_by_xpath("/html/body/div/form[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div/div/input").click()
#     nuvem.find_element_by_xpath("/html/body/div/form/div/div/div[1]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div[1]/input").click()
#     #nuvem.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]/div/div[1]/span/span/button").click()
#     #criar pasta
#     # nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]").click()
#     # nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]").click()
#     nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]/div/div[1]/span/span[1]").click()
#     nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]/div/div[1]/span/span[1]").click()
#     #Criar todas as pastas de cada Solicitação
#     for pastaEspecifica in pastas:
#         nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div/div/div[1]/div[1]/button").click()
#         nuvem.find_element_by_xpath("/html/body/div[7]/div/div/div/div/div/div/ul/li[1]/button").click()
#         nuvem.find_element_by_xpath("/html/body/div[7]/div/div/div/div[2]/div[2]/div/div[2]/div[1]/div/div/div/div/input").send_keys(pastaEspecifica)
#         nuvem.find_element_by_xpath("/html/body/div[3]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div/span/button").click()
#     print("PASTA CRIADA")




    