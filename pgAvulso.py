from time import sleep
from datetime import date
import gerenciadorPastas
from funcoes import padraoChrome
import shutil
import os
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains 
from openpyxl import load_workbook

#retornar a data atual
today = date.today()
#formatar a data com pontos
data_em_texto = today.strftime("%d.%m.%Y")

#caminho da pasta macro(pasta do dia)
caminho_da_pasta = gerenciadorPastas.recuperar_diretorio_usuario() + "\\OneDrive - tpfe.com.br\\RPA-DEV\\" + data_em_texto + "\\" 

#função para tramitar as solicitações
def pagamentoAvulso(financeiro):
    #pra uso de click
    builder = ActionChains(financeiro)
    financeiro.implicitly_wait(40)
    # sleep(10)
    # financeiro.refresh()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/main/section/div/div/div/div/section/div/div[2]/div").click()
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/div[2]/div/div/div[2]/div/div[2]/div[1]").click()
    url = "/runtime/44/list/190/Solicitação de Pgto Avulso"
    #clicar na parte de pagamento avulso
    financeiro.find_element_by_xpath('//a[@href="'+url+'"]').click()

    #Filtro
    #dataA = date.today().strftime("%d/%m/%Y")
    sleep(9)
    #limpar filtro  => financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[2]/button").click()
    filtro = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[3]")
    try:
        filtro_click = builder.click(filtro)
        filtro_click.perform()
    except:
        pass
    try:
        filtro_click = builder.click(filtro)
        filtro_click.perform()
    except:
        pass    
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[3]/div/div/div/div").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[3]/ul/li[3]").send_keys("\n")
    financeiro.find_element_by_xpath("/html/body/div[6]/div[1]").click()
    print(20*"=")
    # financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/ul/li[10]/div/div/div/div/div/div[1]/div/input").send_keys(dataA)
    financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/button").click()#send_keys("\n")
    
    #abrir a linha (Laço para todos itens filtrados)
    sleep(3)
    #pegar o corpo da tabela de solicitaçoes
    tbody1 = financeiro.find_element_by_xpath("//*[@id='mainContent']/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody")
    #pegar as linhas que existem dentro do corpo
    rows1 = tbody1.find_elements_by_tag_name("tr")

    global solicitaçao
    #lista que vai receber os índices de cada linha
    solicitaçao = [] 

    for row in range(len(rows1)):
        solicitaçao.append(row) #armazenando os indices de cada linha

    print(solicitaçao)

    #laço para tramitar cada solicitaçao
    for linha in solicitaçao: 
        # index1 = str(linha+1)
        #para cada linha guardar as informações dela
        global identificador
        #armazenando o id de cada solicitaçao
        identificador = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[4]/div").get_attribute("innerText")
        global razao
        #armazenando a razao social de cada solicitaçao
        razao = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[6]/div").get_attribute("innerText")
        estado = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[7]").get_attribute("value")
        #criando um modelo de nome de pastas para serem salvas(igualmente ao modelo do financeiro)
        #criando uma condicional para saber se a pasta tem apenas id ou tem os dois(id + razao)
        if not razao:
            nomeDaPasta = (f"ID {identificador}")
        else:    
            nomeDaPasta = (f"ID {identificador} {razao}")

        print(nomeDaPasta)
        #com a funçao do outro arquivo, criar a pasta da atual solicitação de acordo com o laço
        gerenciadorPastas.criarPastasFilhas(nomeDaPasta)
        #tempo para salvar todas pastas
        sleep(1.5)
        # #para cada solicitação marcar a caixa de selecionar todas
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        # #desmarcar
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        #clicar na caixa de seleção especifica da linha
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        #clicar no lápis de edição
        financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[1]").click()
        #PEGAR TODAS AS INFORMAÇOES PARA ALIMENTAR A PLANILHA
        cnpj = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[2]/div[1]/div/div/div/input").get_attribute("value")
        banco = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        agencia = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        conta = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[2]/div/div[1]/div/div/div/input").get_attribute("value")
        tipo_conta = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[3]/div[2]/div/div[2]/div/div/div/div/div/div/div[1]/input").get_attribute("value")
        natureza_da_conta = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[3]/div/div/div/div/div/div/div[1]/input").get_attribute("value")
        forma_de_pagamento = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[1]/div/div/div/div/div/div/div[1]/input").get_attribute("value")
        valor = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        valor_pago = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[4]/div[2]/div/div[2]/div/div/div/input").get_attribute("value")
        data_solicitada = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[1]/div/div[1]/div/div/div/input").get_attribute("value")
        data_solicitacao = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[1]/div/div[2]/div/div/div/input").get_attribute("value")
        data_pagamento = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[5]/div[2]/div/div[1]/div/div/div/input").get_attribute("value")


        #clicar em "notas fiscais"
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[2]").click()

        tbody2 = financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[3]/div/div/div/div[1]/div[3]/table/tbody")
        #pega todas as linhas que contem nf
        rows2 = tbody2.find_elements_by_tag_name("a") 
       
        #clicar só se houver elementos 
        try:
            if len(rows2) > 0:
                for row in rows2:
                    row.click()
        except:
            comentarioNf = f"Não foi possível baixar a nota fiscal"            

        sleep(4)    
        #imprimindo
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[1]/div/div[2]/div/button[1]").click() 
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[3]/form/fieldset/div/div/div[2]/div/div[6]/div[2]/div/div[2]/div/div/button").click()  
        #outra forma de clique diretamente no elemento
        financeiro.switch_to_frame(0)
        imprimir = financeiro.find_element_by_xpath("/html/body/div/div/div/div[2]/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/div/table/tbody/tr").click() 
        
        try:
            click_me = builder.click(imprimir)
            click_me.perform()
        except:
            print("1- Não consegui clicar")
        try:
            click_me = builder.click(imprimir)
            click_me.perform()
        except:
            print("2- Não consegui clicar")
            

        financeiro.find_element_by_xpath("/html/body/div/div/div/div[16]/div/div[1]/table/tbody/tr/td[2]").click()
        financeiro.find_element_by_xpath("/html/body/div/div/div/div[20]/div[4]/table/tbody/tr/td[1]/div/table/tbody/tr/td").click()
        financeiro.switch_to.default_content()
        # financeiro.find_element_by_xpath("/html/body/div[8]/div[3]").send_keys(Keys.ESC)
        financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[1]/h2/div/div[2]/button").click()
        #financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[2]").click()
        # financeiro.find_element_by_xpath("/html/body/div[8]/div[3]/div/div[2]/ul/div[3]").click()
        print("passou")
        # financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[4]/fieldset/button[2]").click()
        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div/div/div[1]/div/div[3]/button").click()
        
        arquivo_excel = "C:\\Users\\Usuario\\OneDrive - tpfe.com.br\\RPA-DEV\\Planilha de Acompanhamento de Solicitações Financeiras 2021 (1).xlsx"
        wb = load_workbook(arquivo_excel)
        sh1 = wb.worksheets[0]

        maximo = sh1.max_row
        print(maximo)

        # minimo = lista[0]
        lista = list(range(8,maximo))
        try:
            for x in solicitaçao:
                sh1[f"A{lista[0]}"].value = identificador
                sh1[f"B{lista[0]}"].value = cnpj
                sh1[f"C{lista[0]}"].value = razaoSocial
                sh1[f"D{lista[0]}"].value = forma_de_pagamento
                sh1[f"E{lista[0]}"].value = banco
                sh1[f"F{lista[0]}"].value = agencia
                sh1[f"G{lista[0]}"].value = conta
                sh1[f"H{lista[0]}"].value = tipo_conta
                sh1[f"I{lista[0]}"].value = natureza_da_conta
                sh1[f"J{lista[0]}"].value = valor
                sh1[f"K{lista[0]}"].value = valor_pago
                sh1[f"L{lista[0]}"].value = data_solicitada
                sh1[f"M{lista[0]}"].value = data_solicitacao
                sh1[f"N{lista[0]}"].value = data_pagamento
                sh1[f"Q{lista[0]}"].value = estado
                lista.pop(0)
            print(lista) 
            wb.save(arquivo_excel)
        except:
            print("ERRO AO SALVAR NA PLANILHA")
        sleep(1.5)
        print("mover arquivos")
        #listando os arquivos baixados na pasta macro(pasta do dia)
        arquivos = gerenciadorPastas.listar_arquivos_em_diretorios(caminho_da_pasta)
        print(arquivos)
        #criando um laço para mover cada um para sua pasta especifica
        for arquivo in arquivos:
            print(arquivo)
            #movendo os arquivos para a pasta da sua solicitaçao
            shutil.move(caminho_da_pasta + arquivo, caminho_da_pasta + nomeDaPasta +"\\" + arquivo)
            print("moveu o arquivo!")

        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/tbody/tr[1]/td[2]/span/span[1]/input").click()
        
        tramitar = financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/div[3]/div/button[2]")
        
        try:
            tramitar.send_keys("\n")
        except:
            print("1- Não consegui clicar")
        try:
            tramitar.click()
        except:
            print("1- Não consegui clicar")

        financeiro.find_element_by_xpath("/html/body/div[5]/div[3]/div/div[2]/ul/div[3]").click()
        sleep(1.5)

        # financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[2]").click()    

    #para cada solicitação marcar a caixa de selecionar todas
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    #desmarcar
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[3]/div/div/div/table/thead/tr/th[2]/span/span[1]/input").click()
    #exportando a planilha
    financeiro.find_element_by_xpath("/html/body/div[1]/div/div[2]/div/main/section/div/div/div/div[1]/div/div[1]/button[4]").click()






    