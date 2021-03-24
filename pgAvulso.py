from time import sleep
from datetime import date
import gerenciadorPastas
from main import padraoChrome

today = date.today()
data_em_texto = today.strftime("%d.%m.%Y")

def pgtoAvulso(financeiro):
    financeiro.maximize_window()
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
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[4]/div").get_attribute("innerText")
        global razao
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[6]/div").get_attribute("innerText")
        nomeDaPasta = (f"ID {identificador} {razao}")
        print(nomeDaPasta)
        gerenciadorPastas.criarPastasFilhas(nomeDaPasta)
        sleep(10)
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr["+ index1 +"]/td[2]/span/span[1]/input").click()
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()

        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        rows2 = tbody2.find_elements_by_tag_name("a") # pega todas as linhas que contem nf
        padraoChrome("OneDrive - tpfe.com.br\\RPA-DEV" + data_em_texto + "\\" + nomeDaPasta )
        nf = [] 
        for row in range(len(rows2)):
            nf.append(row)
        print(nf)
        for nota in nf:
        #para cada nf, baixar cada uma delas
            index2 = str(nota+1)
            financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr["+ index2 +"]/td[2]/div[2]/div/a").click()#send_keys("\n")
            print(f"{index}º NF baixada!")
        
        

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




    