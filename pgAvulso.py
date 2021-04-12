from time import sleep
from datetime import date
import gerenciadorPastas
from funcoes import padraoChrome
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 


#retornar o usuario da maquina utilizada pelo robo
# usuario = getpass.getuser()
#retornar a data atual
today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")
pref = gerenciadorPastas.recuperar_diretorio_usuario() + "\\OneDrive - tpfe.com.br\\RPA-DEV\\" + data_em_texto + "\\" 


def pagamentoAvulso(financeiro):
    builder = ActionChains(financeiro)
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
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]").click()
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
    # pastas = []
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
        if not razao:
            nomeDaPasta = (f"ID {identificador}")
        else:    
            nomeDaPasta = (f"ID {identificador} {razao}")
        print(nomeDaPasta)
        #adicionando cada modelo de nome de pasta para uma lista
        # pastas.append(nomeDaPasta)
        #para cada solicitaçao, criar as pastas respectivas
        gerenciadorPastas.criarPastasFilhas(nomeDaPasta)
        #tempo para salvar todas pastas
        sleep(1.5)
        #para cada solicitaçao, clique nelas e baixe as nf e imprima
        #pega a posição dela na lista
        #acrescenta mais um para ser inserido corretamente no xpath
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[2]/span/span[1]/input").click()
        #financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[""]")
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()

        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        rows2 = tbody2.find_elements_by_tag_name("a") 
        # pega todas as linhas que contem nf
        #criaçao de uma lista para armazenar os indices de cada nota fiscal
        #laço para adicionar os valores de *indices* na lista
        if len(rows2) > 0:
            for row in rows2:
                row.click()

        sleep(10)    
        #imprimindo
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click() 
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button").click()  
        #outra forma de clique diretamente no elemento
        financeiro.switch_to_frame(0)
        imprimir = financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr").click() 
        click_me = builder.click(imprimir)
        click_me.perform()

        financeiro.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        financeiro.switch_to.default_content()
        # financeiro.find_element_by_xpath("/html/body/div[8]/div[3]").send_keys(Keys.ESC)
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()
        sleep(5)

        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(pref)
        print(arquivos)
        for arquivo in arquivos:
            print(arquivo)
            shutil.move(pref + arquivo, pref + nomeDaPasta +"\\" + arquivo)
            print("moveu o arquivo!")

    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[4]").click()

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




    