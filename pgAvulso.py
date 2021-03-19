from time import sleep
from datetime import date
def pgtoAvulso(financeiro):
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

    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
    #sleep(2)
    #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[1]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    #nota
    #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
    #baixar

def baixarNf(financeiro):
    financeiro.get("https://tpf.madrix.app/?next=/app")
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

    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
    #sleep(2)
    #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div[1]/div[2]/div/div/div/div/div[1]/div[3]/table/tbody/tr/td[2]/div[2]/div/a").click()
    #nota
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()
    #pegar o id de solicitação + razao social
    global identificador
    identificador = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[2]/span").get_attribute("innerText")
    print(identificador)
    global razao
    razao = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[1]/div/div/input").get_attribute("value")
    print(razao)

    # notaFiscal
    tbody = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
    rows = tbody.find_elements_by_tag_name("a")
    nf = []
    for row in range(len(rows)):
        nf.append(row)
        # notaFiscal += 1
    print(nf)
    for nota in nf:
        index = str(nota+1)
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody/tr["+ index +"]/td[2]/div[2]/div/a").click()#send_keys("\n")
        print(f"{index}º NF baixada!")
    sleep(10)    
def chamarSharepoint(nuvem, Keys):
    
    nuvem.find_element_by_tag_name('body').send_keys(Keys.COMMAND + 't')
    nuvem.get("https://tpfecombr-my.sharepoint.com")
    

    nuvem.find_element_by_xpath("/html/body/div/form[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div/div/div[2]/div[2]/div/input[1]").send_keys("roberta.costa@tpfe.com.br")
    nuvem.find_element_by_xpath("/html/body/div/form[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div/div/div[4]/div/div/div/div/input").click()
    nuvem.find_element_by_xpath("/html/body/div/form[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[2]/div/div[2]/input").send_keys("Tecoteco123")
    nuvem.find_element_by_xpath("/html/body/div/form[1]/div/div/div[2]/div/div/div[1]/div[2]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div/div/input").click()
    nuvem.find_element_by_xpath("/html/body/div/form/div/div/div[1]/div[2]/div/div[2]/div/div[3]/div[2]/div/div/div[1]/input").click()
    #nuvem.find_element_by_xpath("/html/body/div[3]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]/div/div[1]/span/span/button").click()
    #criar pasta
    # nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]").click()
    # nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[2]/main/div/div/div[2]/div/div/div/div/div[2]/div/div/div/div[1]/div/div/div/div[3]").click()
    nuvem.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div/div/div[1]/div[1]/button").click()
    nuvem.find_element_by_xpath("/html/body/div[3]/div/div/div/div/div/div/ul/li[1]/button").click()
    global pastaSolicitacao
    pastaSolicitacao = (f"ID {identificador} {razao}")
    nuvem.find_element_by_xpath("/html/body/div[3]/div/div/div/div[2]/div[2]/div/div[2]/div[1]/div/div/div/div/input").send_keys(pastaSolicitacao)
    nuvem.find_element_by_xpath("/html/body/div[3]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div/span/button").click()
    print("PASTA CRIADA")